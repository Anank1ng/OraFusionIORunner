from __future__ import annotations

import io
import zipfile
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from services.config_loader import all_request_fields, load_schema, schema_to_mapping
from services.excel_service import (
    dictionary_dataframe,
    fields_to_template_dataframe,
    make_json_bytes,
    make_template_excel_bytes,
)
from services.payload_builder import PayloadBuildError, build_payload_from_row
from services.reference_service import is_read_only_post_field, post_safe_mapping

st.set_page_config(page_title="Template & JSON Builder", page_icon="🧱", layout="wide")

PRESET_SECTIONS: Dict[str, List[str]] = {
    "Minimal Create IO": [
        "Core Organization",
        "Financial IDs",
        "Item Definition Settings",
    ],
    "Standard Warehouse IO": [
        "Core Organization",
        "Financial IDs",
        "Item Definition Settings",
        "Additional Usages",
        "Inventory Settings",
        "Movement Request",
        "Picking Defaults",
    ],
    "Full Advanced IO": [
        "Core Organization",
        "Financial IDs",
        "Item Definition Settings",
        "Additional Usages",
        "Inventory Settings",
        "Movement Request",
        "Picking Defaults",
        "Lot Control",
        "Child Lot Control",
        "Serial Number",
        "Item Sourcing Details",
        "Distribution Parameters",
        "Kanban",
        "Packing Unit",
        "Plant Parameters",
    ],
}


def _sanitize_filename(name: str) -> str:
    return "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in name).strip("_") or "template"


def _bundle_zip_bytes(mapping: Dict[str, Any], template_df: pd.DataFrame, sample_payload: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("upload_template.xlsx", make_template_excel_bytes(mapping, template_df=template_df))
        zf.writestr("mapping.json", make_json_bytes(mapping))
        zf.writestr("sample_payload.json", make_json_bytes(sample_payload))
        zf.writestr("field_dictionary.csv", dictionary_dataframe(mapping).to_csv(index=False).encode("utf-8"))
        zf.writestr(
            "readme_upload_guide.txt",
            (
                "Oracle Fusion Inventory Organization Runner\n"
                "1. Isi upload_template.xlsx pada sheet Upload_Template.\n"
                "2. Optional blank cells akan di-skip dari JSON payload.\n"
                "3. Pakai mapping.json saat upload di halaman Upload Runner.\n"
                "4. Cek field_dictionary.csv untuk arti kolom, tipe data, dan referensi ID/LOV.\n"
            ).encode("utf-8"),
        )
    return output.getvalue()


def _field_rows(schema_fields: List[Dict[str, Any]], preset_name: str, post_safe_only: bool) -> pd.DataFrame:
    preset_sections = PRESET_SECTIONS.get(preset_name)
    rows: List[Dict[str, Any]] = []
    for field in schema_fields:
        post_safe = not is_read_only_post_field(field)
        section = field.get("section", "General")
        include = bool(field.get("include_by_default"))
        if preset_sections is not None:
            include = section in preset_sections and post_safe
        if preset_name == "Full Advanced IO":
            include = post_safe
        if post_safe_only and not post_safe:
            continue
        rows.append(
            {
                "include": include,
                "post_safe": post_safe,
                "required": bool(field.get("required")),
                "section": section,
                "label": field.get("label", field.get("excel_column")),
                "excel_column": field.get("excel_column"),
                "payload_path": field.get("payload_path"),
                "type": field.get("type"),
                "default": field.get("default", ""),
                "max_length": field.get("max_length", ""),
                "reference_hint": field.get("reference_hint", ""),
                "use_instead": field.get("use_instead", ""),
                "description": field.get("description", ""),
            }
        )
    return pd.DataFrame(rows)


def _section_badge(field_df: pd.DataFrame, section: str) -> str:
    subset = field_df[field_df["section"] == section]
    selected = int((subset["include"] == True).sum())
    total = len(subset)
    required = int((subset["required"] == True).sum())
    unsafe = int((subset["post_safe"] == False).sum()) if "post_safe" in subset else 0
    return f"{section} · {selected}/{total} field · {required} required" + (f" · {unsafe} display-only" if unsafe else "")


st.title("🧱 Template & JSON Builder")
st.caption("Buat template Excel, mapping JSON, dan sample payload dengan UI bertahap supaya field advanced tidak bikin bingung.")

schema = load_schema("inventory_organizations.json")
fields = all_request_fields(schema)

with st.expander("Informasi API", expanded=False):
    col1, col2, col3 = st.columns(3)
    col1.metric("API", schema.get("api_name", "-"))
    col2.metric("Method", schema.get("method", "-"))
    col3.metric("Version", schema.get("version", "-"))
    st.code(schema.get("endpoint", ""), language="http")
    for note in schema.get("notes", []):
        st.write(f"- {note}")

st.subheader("1. Pilih mode template")
preset = st.radio(
    "Template Mode",
    options=["Minimal Create IO", "Standard Warehouse IO", "Full Advanced IO", "Custom"],
    index=1,
    horizontal=True,
    help="Pilih preset dulu. Setelah itu kamu tetap bisa edit field di section cards.",
)

mode_help = {
    "Minimal Create IO": "Field inti untuk create IO: core, financial ID, dan item definition settings.",
    "Standard Warehouse IO": "Field yang paling sering dipakai untuk warehouse/inventory org, termasuk inventory settings dan picking/movement request.",
    "Full Advanced IO": "Semua field POST-safe dari schema, termasuk lot, serial, kanban, packing, dan plant parameters.",
    "Custom": "Mulai dari default schema lalu pilih manual per section.",
}
st.info(mode_help[preset])

post_safe_only = st.toggle(
    "Tampilkan POST Safe only",
    value=True,
    help="Kalau aktif, field display/read-only seperti LegalEntityName disembunyikan dari pilihan template POST.",
)

field_df = _field_rows(fields, preset, post_safe_only)
if field_df.empty:
    st.error("Tidak ada field yang bisa ditampilkan. Matikan filter POST Safe only untuk melihat field display-only.")
    st.stop()

# Simpan state default antar rerun, tapi reset saat preset/filter berubah.
state_key = f"builder_field_df::{preset}::{post_safe_only}"
if st.session_state.get("builder_active_state_key") != state_key:
    st.session_state["builder_active_state_key"] = state_key
    st.session_state["builder_field_df"] = field_df.copy()

working_df = st.session_state.get("builder_field_df", field_df).copy()

metric_cols = st.columns(5)
metric_cols[0].metric("Total field", len(working_df))
metric_cols[1].metric("Dipilih", int((working_df["include"] == True).sum()))
metric_cols[2].metric("Required", int((working_df["required"] == True).sum()))
metric_cols[3].metric("Section", working_df["section"].nunique())
metric_cols[4].metric("Display-only", int((working_df["post_safe"] == False).sum()))

st.subheader("2. Pilih section dan field")
all_sections = sorted(working_df["section"].dropna().unique().tolist())
default_visible_sections = [s for s in all_sections if int((working_df.loc[working_df["section"] == s, "include"] == True).sum()) > 0]
if not default_visible_sections:
    default_visible_sections = all_sections[:4]
visible_sections = st.multiselect(
    "Section yang ditampilkan",
    options=all_sections,
    default=default_visible_sections,
    help="Ini hanya mengatur tampilan. Field yang sudah dipilih di section lain tetap tersimpan.",
)

c1, c2, c3 = st.columns(3)
if c1.button("✅ Pilih semua field terlihat", use_container_width=True):
    working_df.loc[working_df["section"].isin(visible_sections) & (working_df["post_safe"] == True), "include"] = True
    st.session_state["builder_field_df"] = working_df
    st.rerun()
if c2.button("⬜ Kosongkan field terlihat", use_container_width=True):
    working_df.loc[working_df["section"].isin(visible_sections), "include"] = False
    st.session_state["builder_field_df"] = working_df
    st.rerun()
if c3.button("↩️ Reset sesuai preset", use_container_width=True):
    st.session_state["builder_field_df"] = field_df.copy()
    st.rerun()

for section in visible_sections:
    section_df = working_df[working_df["section"] == section].reset_index(drop=True)
    with st.expander(_section_badge(working_df, section), expanded=bool((section_df["include"] == True).any())):
        edited = st.data_editor(
            section_df,
            hide_index=True,
            use_container_width=True,
            height=min(420, 110 + 36 * max(len(section_df), 1)),
            disabled=[
                "post_safe",
                "required",
                "section",
                "label",
                "excel_column",
                "payload_path",
                "type",
                "max_length",
                "reference_hint",
                "use_instead",
                "description",
            ],
            column_config={
                "include": st.column_config.CheckboxColumn("Include"),
                "post_safe": st.column_config.CheckboxColumn("POST Safe"),
                "required": st.column_config.CheckboxColumn("Required"),
                "label": st.column_config.TextColumn("Label", width="medium"),
                "excel_column": st.column_config.TextColumn("Excel Column", width="medium"),
                "payload_path": st.column_config.TextColumn("Payload Path", width="medium"),
                "type": st.column_config.TextColumn("Type", width="small"),
                "default": st.column_config.TextColumn("Default", width="small"),
                "reference_hint": st.column_config.TextColumn("Reference", width="medium"),
                "use_instead": st.column_config.TextColumn("Use Instead", width="medium"),
                "description": st.column_config.TextColumn("Description", width="large"),
            },
            key=f"builder_editor_{section}_{state_key}",
        )
        for _, row in edited.iterrows():
            mask = working_df["excel_column"] == row["excel_column"]
            working_df.loc[mask, "include"] = bool(row["include"])
            working_df.loc[mask, "default"] = row.get("default", "")

# Guard supaya read-only/display-only tidak pernah masuk mapping POST.
working_df.loc[working_df["post_safe"] == False, "include"] = False
st.session_state["builder_field_df"] = working_df.copy()

selected_columns = working_df.loc[working_df["include"] == True, "excel_column"].tolist()
mapping = post_safe_mapping(schema_to_mapping(schema, selected_columns))

# Apply editable defaults.
default_by_col = dict(zip(working_df["excel_column"], working_df["default"]))
for field in mapping["fields"]:
    default_value = default_by_col.get(field["excel_column"], field.get("default"))
    if default_value == "" or pd.isna(default_value):
        field.pop("default", None)
    else:
        field["default"] = default_value

st.subheader("3. Preview template dan payload")
left, right = st.columns([1.15, 0.85])
with left:
    selected_by_section = (
        working_df[working_df["include"] == True]
        .groupby("section")
        .size()
        .reset_index(name="selected_fields")
        .sort_values("section")
    )
    st.write("Ringkasan section terpilih")
    st.dataframe(selected_by_section, use_container_width=True, height=190)

    template_df = fields_to_template_dataframe(mapping, sample_rows=3)
    st.write("Preview template Excel")
    st.dataframe(template_df, use_container_width=True, height=260)

with right:
    st.write("Preview status")
    status_cols = st.columns(2)
    status_cols[0].metric("Excel columns", len(template_df.columns))
    status_cols[1].metric("Mapping fields", len(mapping.get("fields", [])))

    try:
        sample_payload = build_payload_from_row(template_df.iloc[0], mapping) if not template_df.empty else {}
        st.success("Sample payload valid untuk struktur mapping ini.")
        st.json(sample_payload)
    except PayloadBuildError as exc:
        sample_payload = {}
        st.error(f"Sample payload belum valid: {exc}")

st.subheader("4. Download hasil builder")
file_stem = _sanitize_filename(preset.lower())
dl1, dl2, dl3, dl4 = st.columns(4)
with dl1:
    st.download_button(
        "⬇️ Template Excel",
        data=make_template_excel_bytes(mapping, template_df=template_df),
        file_name=f"inventory_organizations_{file_stem}_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with dl2:
    st.download_button(
        "⬇️ Mapping JSON",
        data=make_json_bytes(mapping),
        file_name=f"inventory_organizations_{file_stem}_mapping.json",
        mime="application/json",
        use_container_width=True,
    )
with dl3:
    st.download_button(
        "⬇️ Sample Payload",
        data=make_json_bytes(sample_payload),
        file_name=f"inventory_organizations_{file_stem}_sample_payload.json",
        mime="application/json",
        use_container_width=True,
    )
with dl4:
    st.download_button(
        "⬇️ Bundle ZIP",
        data=_bundle_zip_bytes(mapping, template_df, sample_payload),
        file_name=f"inventory_organizations_{file_stem}_bundle.zip",
        mime="application/zip",
        use_container_width=True,
    )

with st.expander("Field Dictionary dari mapping final"):
    st.dataframe(dictionary_dataframe(mapping), use_container_width=True)
