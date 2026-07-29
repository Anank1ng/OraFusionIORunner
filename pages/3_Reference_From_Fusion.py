from __future__ import annotations

import io
import zipfile
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from services.config_loader import FULL_ADVANCED_SECTIONS, MINIMAL_CREATE_IO_COLUMNS, STANDARD_WAREHOUSE_SECTIONS, all_request_fields, load_schema, schema_to_mapping
from services.excel_service import make_json_bytes, make_template_excel_bytes
from services.oracle_client import OracleFusionClient
from services.payload_builder import PayloadBuildError, build_payload_from_row
from services.reference_service import (
    REFERENCE_ENDPOINTS,
    available_fields_from_reference,
    fetch_reference_collection,
    flatten_item_to_excel_row,
    item_preview_dataframe,
    items_to_reference_dataframe,
    mapping_with_reference_defaults,
    post_safe_mapping,
    reference_template_dataframe,
    reference_workbook_bytes,
)

st.set_page_config(page_title="Reference Finder", page_icon="🔎", layout="wide")

STANDARD_SECTIONS = STANDARD_WAREHOUSE_SECTIONS
ADVANCED_SECTIONS = FULL_ADVANCED_SECTIONS


def _make_client(base_url: str, username: str, password: str, timeout: int) -> OracleFusionClient:
    return OracleFusionClient(base_url=base_url, username=username, password=password, timeout=int(timeout))


def _endpoint_status(name: str, reference_tables: Dict[str, pd.DataFrame], reference_errors: pd.DataFrame) -> str:
    if name in reference_tables:
        return f"✅ {len(reference_tables[name])} data"
    if isinstance(reference_errors, pd.DataFrame) and not reference_errors.empty:
        if (reference_errors.get("ReferenceType") == name).any():
            return "⚠️ gagal/cek error"
    return "belum fetch"


def _apply_reference_preset(field_df: pd.DataFrame, preset: str) -> pd.DataFrame:
    df = field_df.copy()
    if df.empty:
        return df
    if preset == "Minimal Create IO":
        df["include"] = df["post_safe"] & df["excel_column"].isin(MINIMAL_CREATE_IO_COLUMNS)
    elif preset == "Standard Warehouse IO":
        df["include"] = df["post_safe"] & df["section"].isin(STANDARD_SECTIONS)
    elif preset == "Copy selected IO settings":
        df["include"] = df["post_safe"] & (
            (df["required"] == True)
            | (df["reference_value"].astype(str).str.strip() != "")
            | (df["default"].astype(str).str.strip() != "")
        )
    elif preset == "Full Advanced Template":
        df["include"] = df["post_safe"]
    # Custom keeps the existing state.
    df.loc[df["post_safe"] == False, "include"] = False
    return df


def _section_badge(field_df: pd.DataFrame, section: str) -> str:
    subset = field_df[field_df["section"] == section]
    selected = int((subset["include"] == True).sum())
    total = len(subset)
    with_value = int((subset["reference_value"].astype(str).str.strip() != "").sum()) if "reference_value" in subset else 0
    unsafe = int((subset["post_safe"] == False).sum()) if "post_safe" in subset else 0
    return f"{section} · {selected}/{total} selected · {with_value} punya value" + (f" · {unsafe} display-only" if unsafe else "")


def _bundle_zip_bytes(
    reference_mapping: Dict[str, Any],
    template_df: pd.DataFrame,
    sample_payload: Dict[str, Any],
    reference_tables: Dict[str, pd.DataFrame],
    reference_errors: pd.DataFrame,
    response_state: Dict[str, Any],
    selected_item: Dict[str, Any],
) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("upload_template_from_reference.xlsx", make_template_excel_bytes(reference_mapping, template_df=template_df))
        zf.writestr("mapping_from_reference.json", make_json_bytes(reference_mapping))
        zf.writestr("sample_payload_from_reference.json", make_json_bytes(sample_payload))
        zf.writestr("selected_reference_raw.json", make_json_bytes(selected_item))
        zf.writestr("get_existing_org_response_debug.json", make_json_bytes(response_state))
        if reference_tables or (isinstance(reference_errors, pd.DataFrame) and not reference_errors.empty):
            zf.writestr("reference_ids.xlsx", reference_workbook_bytes(reference_tables, reference_errors))
        zf.writestr(
            "readme_reference_upload_guide.txt",
            (
                "Oracle Fusion Inventory Organization Runner - Reference Bundle\n"
                "1. Gunakan reference_ids.xlsx untuk mencari ID/LOV seperti BusinessUnitId, LegalEntityId, ScheduleId, LocationId.\n"
                "2. Isi upload_template_from_reference.xlsx pada sheet Upload_Template.\n"
                "3. Gunakan mapping_from_reference.json saat upload di halaman Upload Runner.\n"
                "4. Optional blank cells akan di-skip dari JSON payload.\n"
                "5. Field display-only/name tidak dimasukkan ke mapping POST. Gunakan field ID/Code penggantinya.\n"
            ).encode("utf-8"),
        )
    return output.getvalue()


st.title("🔎 Reference Finder")
st.caption("Version: v12.1 field selection fix — manual checkbox ikut preview/download")
st.caption("Ambil ID referensi dari Oracle, pilih existing IO sebagai contoh, lalu generate template upload yang POST-safe.")

schema = load_schema("inventory_organizations.json")
default_mapping = schema_to_mapping(schema)
full_reference_mapping = schema_to_mapping(schema, [field["excel_column"] for field in all_request_fields(schema)])

try:
    default_base_url = st.secrets.get("ORACLE_BASE_URL", "")
    default_username = st.secrets.get("ORACLE_USERNAME", "")
    default_password = st.secrets.get("ORACLE_PASSWORD", "")
except Exception:
    default_base_url = ""
    default_username = ""
    default_password = ""

st.sidebar.header("Oracle Connection")
base_url = st.sidebar.text_input("Oracle Base URL", value=default_base_url, placeholder="https://your-instance.oraclecloud.com")
username = st.sidebar.text_input("Username", value=default_username)
password = st.sidebar.text_input("Password", value=default_password, type="password")
timeout = st.sidebar.number_input("Timeout per request (seconds)", min_value=10, max_value=300, value=60, step=10)
connection_disabled = not base_url or not username or not password

if connection_disabled:
    st.info("Isi Oracle Base URL, username, dan password di sidebar dulu.")

# -----------------------------------------------------------------------------
# 1. LOV / master reference fetcher
# -----------------------------------------------------------------------------
st.subheader("1. Ambil Reference Data / LOV")
st.write("Pakai ini untuk mencari ID seperti BusinessUnitId, ProfitCenterBusinessUnitId, LegalEntityId, OrganizationId, ScheduleId, dan LocationId.")

reference_tables = st.session_state.get("lov_reference_tables", {})
reference_errors = st.session_state.get("lov_reference_errors", pd.DataFrame())
raw_lov = st.session_state.get("lov_raw_responses", {})

with st.container(border=True):
    st.write("**Endpoint reference yang akan diambil**")
    selected_reference_names: List[str] = []
    cols = st.columns(2)
    for idx, (name, config) in enumerate(REFERENCE_ENDPOINTS.items()):
        with cols[idx % 2]:
            default_checked = not bool(config.get("optional"))
            checked = st.checkbox(name, value=default_checked, key=f"lov_{name}")
            if checked:
                selected_reference_names.append(name)
            st.caption(config["endpoint"])
            if config.get("fixed_q_filter"):
                st.caption(f"fixed q: {config['fixed_q_filter']}")
            if config.get("description"):
                st.caption(config["description"])
            st.caption(_endpoint_status(name, reference_tables, reference_errors))

lov_col1, lov_col2 = st.columns([1, 2])
with lov_col1:
    lov_limit = st.number_input("Limit per reference endpoint", min_value=1, max_value=500, value=100, step=25)
with lov_col2:
    lov_q_filter = st.text_input(
        "q filter untuk LOV opsional",
        value="",
        placeholder="Name LIKE 'Vision%'",
        help="Dikosongkan saja kalau mau ambil daftar umum. Filter ini dikirim ke semua endpoint yang dipilih.",
    )

if st.button("📥 Fetch Selected Reference", disabled=connection_disabled or not selected_reference_names, use_container_width=True):
    client = _make_client(base_url, username, password, int(timeout))
    reference_tables = {}
    raw_responses: Dict[str, Any] = {}
    errors: List[Dict[str, Any]] = []
    progress = st.progress(0)

    with st.spinner("Mengambil reference data dari Oracle..."):
        for idx, name in enumerate(selected_reference_names):
            config = REFERENCE_ENDPOINTS[name]
            try:
                response = fetch_reference_collection(
                    client,
                    endpoint=config["endpoint"],
                    limit=int(lov_limit),
                    q_filter=lov_q_filter,
                    fixed_q_filter=config.get("fixed_q_filter", ""),
                )
                raw_responses[name] = {"ok": response.ok, "status_code": response.status_code, "url": response.url, "body": response.body}
                if not response.ok:
                    errors.append({
                        "ReferenceType": name,
                        "endpoint": config["endpoint"],
                        "status_code": response.status_code,
                        "message": str(response.body)[:1500],
                    })
                    continue
                body = response.body
                items = body.get("items", []) if isinstance(body, dict) else []
                if not isinstance(items, list):
                    items = []
                reference_tables[name] = items_to_reference_dataframe(name, items, config)
            except Exception as exc:
                errors.append({
                    "ReferenceType": name,
                    "endpoint": config.get("endpoint", ""),
                    "status_code": "ERROR",
                    "message": str(exc),
                })
            progress.progress((idx + 1) / max(len(selected_reference_names), 1))

    st.session_state["lov_reference_tables"] = reference_tables
    st.session_state["lov_reference_errors"] = pd.DataFrame(errors)
    st.session_state["lov_raw_responses"] = raw_responses
    st.rerun()

reference_tables = st.session_state.get("lov_reference_tables", {})
reference_errors = st.session_state.get("lov_reference_errors", pd.DataFrame())
raw_lov = st.session_state.get("lov_raw_responses", {})

if reference_tables or (isinstance(reference_errors, pd.DataFrame) and not reference_errors.empty):
    st.write("**Hasil Reference IDs**")
    if reference_tables:
        tabs = st.tabs(list(reference_tables.keys()))
        for tab, (name, df_ref) in zip(tabs, reference_tables.items()):
            with tab:
                if df_ref.empty:
                    st.warning("Tidak ada data di response endpoint ini.")
                else:
                    st.dataframe(df_ref, use_container_width=True, height=240)

    if isinstance(reference_errors, pd.DataFrame) and not reference_errors.empty:
        with st.expander("Endpoint yang gagal / tidak tersedia", expanded=True):
            st.dataframe(reference_errors, use_container_width=True)

    dl_ref1, dl_ref2 = st.columns(2)
    with dl_ref1:
        st.download_button(
            "⬇️ Download Reference IDs Excel",
            data=reference_workbook_bytes(reference_tables, reference_errors),
            file_name="oracle_fusion_reference_ids.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl_ref2:
        st.download_button(
            "⬇️ Download Raw Reference JSON",
            data=make_json_bytes(raw_lov),
            file_name="oracle_fusion_reference_raw.json",
            mime="application/json",
            use_container_width=True,
        )

st.divider()

# -----------------------------------------------------------------------------
# 2. Existing Inventory Organization reference template
# -----------------------------------------------------------------------------
st.sidebar.header("Existing Org GET Options")
limit = st.sidebar.number_input("Existing org limit", min_value=1, max_value=500, value=25, step=5)
offset = st.sidebar.number_input("Existing org offset", min_value=0, max_value=100000, value=0, step=25)
expand_children = st.sidebar.toggle(
    "Expand child params",
    value=True,
    help="Mencoba mengambil child data invOrgParameters dan plantParameters. Kalau instance tidak support, matikan toggle ini.",
)
q_filter = st.sidebar.text_input("Existing org q filter opsional", value="", placeholder="OrganizationCode='M1'")

st.subheader("2. Ambil Existing Inventory Organization")
st.write("Pilih satu organization yang sudah ada sebagai reference. App akan membaca core field + child parameter jika expand aktif.")
endpoint = default_mapping["endpoint"]
with st.container(border=True):
    col_api, col_button = st.columns([2, 1])
    with col_api:
        st.write("Endpoint GET existing org")
        st.code(f"GET {endpoint}", language="http")
        if base_url:
            st.caption(f"Target instance: {base_url.rstrip('/')}")
    with col_button:
        get_clicked = st.button("🔎 GET Existing Orgs", disabled=connection_disabled, use_container_width=True)

if get_clicked:
    params: Dict[str, Any] = {"limit": int(limit), "offset": int(offset), "totalResults": "true"}
    if expand_children:
        params["expand"] = "invOrgParameters,plantParameters"
    if q_filter.strip():
        params["q"] = q_filter.strip()
    with st.spinner("Mengambil existing Inventory Organizations dari Oracle Fusion..."):
        try:
            response = _make_client(base_url, username, password, int(timeout)).get_collection_items(endpoint, params=params)
        except Exception as exc:
            st.error(f"GET error: {exc}")
            st.stop()
    st.session_state["reference_response"] = {
        "ok": response.ok,
        "status_code": response.status_code,
        "body": response.body,
        "url": response.url,
        "params": params,
    }
    # Reset pilihan reference supaya tidak pakai state lama.
    for key in ["reference_field_df", "reference_active_state_key"]:
        st.session_state.pop(key, None)

response_state = st.session_state.get("reference_response")
if not response_state:
    st.info("Klik **GET Existing Orgs** kalau mau generate template dari organization yang sudah ada.")
    st.stop()

if response_state["ok"]:
    st.success(f"GET existing org berhasil. Status {response_state['status_code']}")
else:
    st.error(f"GET existing org gagal. Status {response_state['status_code']}")

with st.expander("Debug response existing org", expanded=not response_state["ok"]):
    st.write("Final URL")
    st.code(response_state.get("url", ""), language="text")
    st.json(response_state.get("body", {}))

body = response_state.get("body")
if not isinstance(body, dict):
    st.error("Response Oracle bukan JSON object, jadi belum bisa diproses sebagai collection.")
    st.stop()
items = body.get("items")
if not isinstance(items, list):
    st.error("Response tidak punya key 'items'. Cek Debug response untuk melihat formatnya.")
    st.stop()
if not items:
    st.warning("Tidak ada Inventory Organization yang ditemukan dari GET ini. Coba ubah limit/filter/expand.")
    st.stop()

# -----------------------------------------------------------------------------
# 3. Choose existing org
# -----------------------------------------------------------------------------
st.subheader("3. Pilih existing organization sebagai reference")
preview_df = item_preview_dataframe(items)
search_text = st.text_input("Search code/name/status", value="", placeholder="Contoh: PAM, M1, Active")
filtered_preview = preview_df.copy()
if search_text.strip():
    query = search_text.strip().lower()
    mask = filtered_preview.astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)
    filtered_preview = filtered_preview[mask]

if filtered_preview.empty:
    st.warning("Tidak ada hasil yang cocok dengan search. Hapus/ubah keyword search.")
    st.stop()

label_by_index: Dict[int, str] = {}
for _, row in filtered_preview.iterrows():
    idx = int(row["result_index"])
    code = row.get("OrganizationCode", "")
    name = row.get("OrganizationName", "")
    status = row.get("Status", "")
    org_id = row.get("OrganizationId", "")
    label_by_index[idx] = f"{idx} · {code} · {name} · {status} · ID {org_id}"

selected_index = st.selectbox(
    "Pilih organization",
    options=list(label_by_index.keys()),
    format_func=lambda idx: label_by_index.get(idx, str(idx)),
)
selected_item = items[int(selected_index)]

st.dataframe(filtered_preview, use_container_width=True, height=220)
with st.expander("Raw JSON reference terpilih", expanded=False):
    st.json(selected_item)

reference_row = flatten_item_to_excel_row(selected_item, full_reference_mapping)
reference_df = pd.DataFrame([reference_row])

summary_items = [
    ("Organization", reference_row.get("OrganizationCode", ""), reference_row.get("OrganizationName", "")),
    ("Management BU", reference_row.get("ManagementBusinessUnitId", ""), reference_row.get("ManagementBusinessUnitName", "")),
    ("Legal Entity", reference_row.get("LegalEntityId", ""), reference_row.get("LegalEntityName", "")),
    ("Profit Center BU", reference_row.get("ProfitCenterBusinessUnitId", ""), reference_row.get("ProfitCenterBusinessUnitName", "")),
    ("Master Org", reference_row.get("MasterOrganizationId", ""), reference_row.get("MasterOrganizationCode", "")),
    ("Item Definition", reference_row.get("ItemDefinitionOrganizationCode", ""), reference_row.get("ItemDefinitionOrganizationName", "")),
    ("Schedule", reference_row.get("invOrgParameters.ScheduleId", ""), reference_row.get("invOrgParameters.ScheduleName", "")),
    ("Location", reference_row.get("LocationId", ""), reference_row.get("LocationName", "")),
]
summary_df = pd.DataFrame(summary_items, columns=["Area", "ID/Code", "Name/Description"])
st.write("Reference summary")
st.dataframe(summary_df, use_container_width=True, height=260)

with st.expander("Semua nilai reference yang cocok dengan schema/template", expanded=False):
    st.dataframe(reference_df, use_container_width=True)

# -----------------------------------------------------------------------------
# 4. Field selection + preview
# -----------------------------------------------------------------------------
st.subheader("4. Generate template dari reference")
st.write("Pilih preset dulu, lalu edit field per section. Field display-only/name tetap terlihat sebagai referensi, tapi otomatis tidak masuk mapping POST.")

base_field_rows = available_fields_from_reference(full_reference_mapping, reference_row)
base_field_df = pd.DataFrame(base_field_rows)

preset = st.radio(
    "Preset",
    options=["Minimal Create IO", "Standard Warehouse IO", "Copy selected IO settings", "Full Advanced Template", "Custom"],
    index=2,
    horizontal=True,
)

state_key = f"reference_fields::{selected_index}::{preset}"
if st.session_state.get("reference_active_state_key") != state_key:
    st.session_state["reference_active_state_key"] = state_key
    st.session_state["reference_field_df"] = _apply_reference_preset(base_field_df, preset)

field_df = st.session_state.get("reference_field_df", _apply_reference_preset(base_field_df, preset)).copy()

unsafe_count = int((field_df["post_safe"] == False).sum()) if "post_safe" in field_df.columns else 0
if unsafe_count:
    st.info(
        f"{unsafe_count} field display/read-only otomatis tidak di-include untuk template POST. "
        "Gunakan field ID/Code seperti LegalEntityId, ProfitCenterBusinessUnitId, dan ItemDefinitionOrganizationCode."
    )

metrics = st.columns(5)
metrics[0].metric("Field terpilih", int((field_df["include"] == True).sum()))
metrics[1].metric("POST Safe", int((field_df["post_safe"] == True).sum()))
metrics[2].metric("Display-only", unsafe_count)
metrics[3].metric("Dengan reference value", int((field_df["reference_value"].astype(str).str.strip() != "").sum()))
metrics[4].metric("Section", field_df["section"].nunique())

all_sections = sorted(field_df["section"].dropna().unique().tolist())
default_sections = [section for section in all_sections if int((field_df.loc[field_df["section"] == section, "include"] == True).sum()) > 0]
if not default_sections:
    default_sections = all_sections[:4]
selected_sections = st.multiselect(
    "Section yang ditampilkan",
    options=all_sections,
    default=default_sections,
    help="Ini hanya filter tampilan. Field dari section lain yang sudah include tetap tersimpan.",
)

btn1, btn2, btn3 = st.columns(3)
if btn1.button("✅ Pilih semua field terlihat", use_container_width=True):
    field_df.loc[field_df["section"].isin(selected_sections) & (field_df["post_safe"] == True), "include"] = True
    st.session_state["reference_field_df"] = field_df
    st.rerun()
if btn2.button("⬜ Kosongkan field terlihat", use_container_width=True):
    field_df.loc[field_df["section"].isin(selected_sections), "include"] = False
    st.session_state["reference_field_df"] = field_df
    st.rerun()
if btn3.button("↩️ Reset preset", use_container_width=True):
    st.session_state["reference_field_df"] = _apply_reference_preset(base_field_df, preset)
    st.rerun()

left_select, right_preview = st.columns([1.05, 0.95])
with left_select:
    for section in selected_sections:
        section_df = field_df[field_df["section"] == section].reset_index(drop=True)
        with st.expander(_section_badge(field_df, section), expanded=bool((section_df["include"] == True).any())):
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
                    "reference_value",
                    "use_instead",
                    "reference_hint",
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
                    "reference_value": st.column_config.TextColumn("Reference Value", width="medium"),
                    "use_instead": st.column_config.TextColumn("Use Instead", width="medium"),
                    "reference_hint": st.column_config.TextColumn("Reference", width="medium"),
                    "description": st.column_config.TextColumn("Description", width="large"),
                },
                key=f"reference_editor_{section}_{state_key}",
            )
            for _, row in edited.iterrows():
                mask = field_df["excel_column"] == row["excel_column"]
                field_df.loc[mask, "include"] = bool(row["include"])

field_df.loc[field_df["post_safe"] == False, "include"] = False
st.session_state["reference_field_df"] = field_df.copy()

selected_columns_raw = field_df.loc[(field_df["include"] == True) & (field_df["post_safe"] == True), "excel_column"].tolist()
# Jangan paksa kembali ke 12 field Minimal setelah user centang manual.
# Minimal hanya dipakai sebagai default awal; setelah itu checkbox user menjadi source utama.
# Tetap rapikan urutan: field minimal dulu, lalu field tambahan sesuai urutan schema.
selected_columns = [col for col in MINIMAL_CREATE_IO_COLUMNS if col in selected_columns_raw]
selected_columns += [col for col in selected_columns_raw if col not in selected_columns]
reference_mapping = post_safe_mapping(schema_to_mapping(schema, selected_columns))

opt1, opt2 = st.columns(2)
with opt1:
    use_reference_defaults = st.toggle(
        "Simpan reference value sebagai default di mapping JSON",
        value=True,
        help="Kalau aktif, field kosong di Excel bisa memakai default dari reference. Field unik tetap bisa dikosongkan.",
    )
with opt2:
    blank_unique_fields = st.toggle(
        "Kosongkan field unik untuk data baru",
        value=True,
        help="OrganizationCode dan OrganizationName akan dibuat placeholder supaya tidak menduplikasi reference.",
    )

reference_mapping = mapping_with_reference_defaults(
    reference_mapping,
    reference_row,
    use_reference_defaults=use_reference_defaults,
    blank_unique_fields=blank_unique_fields,
)

template_df = reference_template_dataframe(
    reference_mapping,
    reference_row,
    sample_rows=3,
    blank_unique_fields=blank_unique_fields,
)

with right_preview:
    st.write("Preview hasil")
    preview_metrics = st.columns(3)
    preview_metrics[0].metric("Excel columns", len(template_df.columns))
    preview_metrics[1].metric("Mapping fields", len(reference_mapping.get("fields", [])))
    preview_metrics[2].metric("POST Safe", "Yes ✅")

    selected_by_section = (
        field_df[field_df["include"] == True]
        .groupby("section")
        .size()
        .reset_index(name="selected_fields")
        .sort_values("section")
    )
    st.dataframe(selected_by_section, use_container_width=True, height=160)

    try:
        sample_payload = build_payload_from_row(template_df.iloc[0], reference_mapping) if not template_df.empty else {}
        st.success("Sample payload valid dari row pertama template.")
        st.json(sample_payload)
    except PayloadBuildError as exc:
        sample_payload = {}
        st.error(f"Sample payload belum valid: {exc}")

st.write("Preview template dari reference")
st.dataframe(template_df, use_container_width=True, height=260)

st.subheader("5. Download template dari reference")
dl1, dl2, dl3, dl4, dl5 = st.columns(5)
with dl1:
    st.download_button(
        "⬇️ Template Excel",
        data=make_template_excel_bytes(reference_mapping, template_df=template_df),
        file_name="inventory_org_template_from_fusion_reference.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with dl2:
    st.download_button(
        "⬇️ Mapping JSON",
        data=make_json_bytes(reference_mapping),
        file_name="inventory_org_mapping_from_fusion_reference.json",
        mime="application/json",
        use_container_width=True,
    )
with dl3:
    st.download_button(
        "⬇️ Sample Payload",
        data=make_json_bytes(sample_payload),
        file_name="inventory_org_sample_payload_from_reference.json",
        mime="application/json",
        use_container_width=True,
    )
with dl4:
    st.download_button(
        "⬇️ Raw GET JSON",
        data=make_json_bytes(response_state),
        file_name="inventory_org_get_response_debug.json",
        mime="application/json",
        use_container_width=True,
    )
with dl5:
    st.download_button(
        "⬇️ Bundle ZIP",
        data=_bundle_zip_bytes(reference_mapping, template_df, sample_payload, reference_tables, reference_errors, response_state, selected_item),
        file_name="inventory_org_reference_upload_bundle.zip",
        mime="application/zip",
        use_container_width=True,
    )

st.info("Setelah template dan mapping JSON didownload, buka halaman Upload Runner. Upload mapping JSON tersebut, lalu upload Excel yang sudah kamu isi.")
