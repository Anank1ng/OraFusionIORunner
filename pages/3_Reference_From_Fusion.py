from __future__ import annotations

from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from services.config_loader import all_request_fields, load_schema, schema_to_mapping
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
    reference_template_dataframe,
    reference_workbook_bytes,
)

st.set_page_config(page_title="Reference Finder", page_icon="🔎", layout="wide")

st.title("🔎 Reference Finder")
st.caption("Ambil ID referensi dari instance Oracle, lalu buat template upload yang lebih aman untuk POST.")

schema = load_schema("inventory_organizations.json")
default_mapping = schema_to_mapping(schema)
full_reference_mapping = schema_to_mapping(
    schema,
    [field["excel_column"] for field in all_request_fields(schema)]
)

try:
    default_base_url = st.secrets.get("ORACLE_BASE_URL", "")
    default_username = st.secrets.get("ORACLE_USERNAME", "")
    default_password = st.secrets.get("ORACLE_PASSWORD", "")
except Exception:
    default_base_url = ""
    default_username = ""
    default_password = ""

st.sidebar.header("Oracle Connection")
base_url = st.sidebar.text_input(
    "Oracle Base URL",
    value=default_base_url,
    placeholder="https://your-instance.oraclecloud.com",
)
username = st.sidebar.text_input("Username", value=default_username)
password = st.sidebar.text_input("Password", value=default_password, type="password")
timeout = st.sidebar.number_input("Timeout per request (seconds)", min_value=10, max_value=300, value=60, step=10)


def make_client() -> OracleFusionClient:
    return OracleFusionClient(base_url=base_url, username=username, password=password, timeout=int(timeout))


connection_disabled = not base_url or not username or not password
if connection_disabled:
    st.info("Isi Oracle Base URL, username, dan password di sidebar dulu.")

# -----------------------------------------------------------------------------
# 1. LOV / master reference fetcher
# -----------------------------------------------------------------------------
st.subheader("1. Ambil daftar ID dari Oracle")
st.write("Pakai ini untuk mencari ID seperti BusinessUnitId, LegalEntityId, OrganizationId, ScheduleId, dan LocationId.")

with st.expander("Endpoint reference yang akan diambil", expanded=True):
    selected_reference_names: List[str] = []
    cols = st.columns(2)
    for idx, (name, config) in enumerate(REFERENCE_ENDPOINTS.items()):
        with cols[idx % 2]:
            default_checked = not bool(config.get("optional"))
            if st.checkbox(name, value=default_checked, key=f"lov_{name}"):
                selected_reference_names.append(name)
            st.caption(config["endpoint"])

lov_limit = st.number_input("Limit per reference endpoint", min_value=1, max_value=500, value=100, step=25)
lov_q_filter = st.text_input(
    "q filter untuk LOV opsional",
    value="",
    placeholder="Name LIKE 'Vision%'",
    help="Dikosongkan saja kalau mau ambil daftar umum. Filter ini dikirim ke semua endpoint yang dipilih.",
)

if st.button("📥 Fetch Reference IDs", disabled=connection_disabled or not selected_reference_names):
    client = make_client()
    reference_tables: Dict[str, pd.DataFrame] = {}
    raw_responses: Dict[str, Any] = {}
    errors: List[Dict[str, Any]] = []

    with st.spinner("Mengambil reference data dari Oracle..."):
        for name in selected_reference_names:
            config = REFERENCE_ENDPOINTS[name]
            try:
                response = fetch_reference_collection(
                    client,
                    endpoint=config["endpoint"],
                    limit=int(lov_limit),
                    q_filter=lov_q_filter,
                )
                raw_responses[name] = {
                    "ok": response.ok,
                    "status_code": response.status_code,
                    "url": response.url,
                    "body": response.body,
                }
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

    st.session_state["lov_reference_tables"] = reference_tables
    st.session_state["lov_reference_errors"] = pd.DataFrame(errors)
    st.session_state["lov_raw_responses"] = raw_responses

reference_tables = st.session_state.get("lov_reference_tables", {})
reference_errors = st.session_state.get("lov_reference_errors", pd.DataFrame())
raw_lov = st.session_state.get("lov_raw_responses", {})

if reference_tables or (isinstance(reference_errors, pd.DataFrame) and not reference_errors.empty):
    st.write("Hasil Reference IDs")
    if reference_tables:
        tabs = st.tabs(list(reference_tables.keys()))
        for tab, (name, df_ref) in zip(tabs, reference_tables.items()):
            with tab:
                if df_ref.empty:
                    st.warning("Tidak ada data di response endpoint ini.")
                else:
                    st.dataframe(df_ref, use_container_width=True, height=260)

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
        )
    with dl_ref2:
        st.download_button(
            "⬇️ Download Raw Reference JSON",
            data=make_json_bytes(raw_lov),
            file_name="oracle_fusion_reference_raw.json",
            mime="application/json",
        )

st.divider()

# -----------------------------------------------------------------------------
# 2. Existing Inventory Organization reference template
# -----------------------------------------------------------------------------
st.subheader("2. Ambil existing Inventory Organization sebagai contoh template")
st.write("Bagian ini mengambil contoh org yang sudah ada, termasuk child invOrgParameters/plantParameters kalau expand aktif, lalu mengubahnya menjadi template upload berbasis field ID dan optional settings.")

st.sidebar.header("Existing Org GET Options")
limit = st.sidebar.number_input("Existing org limit", min_value=1, max_value=500, value=25, step=5)
offset = st.sidebar.number_input("Existing org offset", min_value=0, max_value=100000, value=0, step=25)
expand_children = st.sidebar.toggle(
    "Expand child params",
    value=True,
    help="Mencoba mengambil child data invOrgParameters dan plantParameters. Kalau instance tidak support, matikan toggle ini.",
)
q_filter = st.sidebar.text_input(
    "Existing org q filter opsional",
    value="",
    placeholder="OrganizationCode='M1'",
)

endpoint = default_mapping["endpoint"]
col_api, col_button = st.columns([2, 1])
with col_api:
    st.write("Endpoint GET existing org")
    st.code(f"GET {endpoint}", language="http")
    if base_url:
        st.write("Target instance")
        st.code(base_url.rstrip("/"), language="text")

with col_button:
    get_clicked = st.button("🔎 GET Existing Orgs", disabled=connection_disabled, use_container_width=True)

if get_clicked:
    params: Dict[str, Any] = {
        "limit": int(limit),
        "offset": int(offset),
        "totalResults": "true",
    }
    if expand_children:
        params["expand"] = "invOrgParameters,plantParameters"
    if q_filter.strip():
        params["q"] = q_filter.strip()

    with st.spinner("Mengambil existing Inventory Organizations dari Oracle Fusion..."):
        try:
            response = make_client().get_collection_items(endpoint, params=params)
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

st.subheader("3. Pilih existing organization sebagai reference")
preview_df = item_preview_dataframe(items)
st.dataframe(preview_df, use_container_width=True, height=260)

selected_index = st.number_input(
    "Pilih result_index sebagai reference",
    min_value=0,
    max_value=len(items) - 1,
    value=0,
    step=1,
)
selected_item = items[int(selected_index)]

with st.expander("Raw JSON reference terpilih", expanded=False):
    st.json(selected_item)

reference_row = flatten_item_to_excel_row(selected_item, full_reference_mapping)
reference_df = pd.DataFrame([reference_row])
st.write("Nilai reference yang cocok dengan seluruh field schema/template")
st.dataframe(reference_df, use_container_width=True)

st.subheader("4. Pilih field dan generate template")
st.write("Default field sekarang memakai ID field. Hindari field Name karena biasanya read-only untuk POST.")

field_rows = available_fields_from_reference(full_reference_mapping, reference_row)
field_df = pd.DataFrame(field_rows)

all_sections = sorted(field_df["section"].dropna().unique().tolist())
default_sections = [
    section for section in all_sections
    if section in [
        "Core Organization",
        "Financial IDs",
        "Item Definition Settings",
        "Additional Usages",
        "Inventory Settings",
        "Movement Request",
        "Picking Defaults",
        "Lot Control",
        "Serial Number",
        "Item Sourcing Details",
        "Distribution Parameters",
        "Kanban",
        "Packing Unit",
    ]
]
selected_sections = st.multiselect(
    "Filter section",
    options=all_sections,
    default=default_sections or all_sections,
    help="Pakai filter ini supaya field reference yang tampil tidak terlalu panjang."
)
visible_df = field_df[field_df["section"].isin(selected_sections)].reset_index(drop=True)

edited_visible_df = st.data_editor(
    visible_df,
    hide_index=True,
    use_container_width=True,
    height=520,
    disabled=["required", "section", "label", "excel_column", "payload_path", "type", "reference_value", "reference_hint", "description"],
    column_config={
        "include": st.column_config.CheckboxColumn("Include"),
        "required": st.column_config.CheckboxColumn("Required"),
        "section": st.column_config.TextColumn("Section", width="medium"),
        "label": st.column_config.TextColumn("Label", width="medium"),
        "reference_value": st.column_config.TextColumn("Reference Value", width="large"),
        "reference_hint": st.column_config.TextColumn("Reference", width="medium"),
        "description": st.column_config.TextColumn("Description", width="large"),
    },
)

# Merge edited visible rows back to full list, preserving previously included fields from hidden sections.
edited_df = field_df.copy()
for _, edited_row in edited_visible_df.iterrows():
    mask = edited_df["excel_column"] == edited_row["excel_column"]
    edited_df.loc[mask, "include"] = edited_row["include"]

selected_columns = edited_df.loc[edited_df["include"] == True, "excel_column"].tolist()
reference_mapping = schema_to_mapping(schema, selected_columns)

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
        help="OrganizationCode dan OrganizationName akan dibuat placeholder supaya tidak ikut menduplikasi reference.",
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

col_template, col_payload = st.columns([1.2, 1])
with col_template:
    st.write("Preview template dari reference")
    st.dataframe(template_df, use_container_width=True)

with col_payload:
    try:
        sample_payload = build_payload_from_row(template_df.iloc[0], reference_mapping)
        st.write("Sample payload dari row pertama template")
        st.json(sample_payload)
    except PayloadBuildError as exc:
        sample_payload = {}
        st.error(f"Sample payload belum valid: {exc}")

st.subheader("5. Download template dari reference")
dl1, dl2, dl3, dl4 = st.columns(4)
with dl1:
    st.download_button(
        "⬇️ Template Excel",
        data=make_template_excel_bytes(reference_mapping, template_df=template_df),
        file_name="inventory_org_template_from_fusion_reference.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with dl2:
    st.download_button(
        "⬇️ Mapping JSON",
        data=make_json_bytes(reference_mapping),
        file_name="inventory_org_mapping_from_fusion_reference.json",
        mime="application/json",
    )
with dl3:
    st.download_button(
        "⬇️ Sample Payload JSON",
        data=make_json_bytes(sample_payload),
        file_name="inventory_org_sample_payload_from_reference.json",
        mime="application/json",
    )
with dl4:
    st.download_button(
        "⬇️ Raw GET Response JSON",
        data=make_json_bytes(response_state),
        file_name="inventory_org_get_response_debug.json",
        mime="application/json",
    )

st.info("Setelah template dan mapping JSON didownload, buka halaman Upload Runner. Upload mapping JSON tersebut, lalu upload Excel yang sudah kamu isi.")
