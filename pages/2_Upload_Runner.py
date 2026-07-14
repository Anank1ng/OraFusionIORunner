from __future__ import annotations

import io
import json
from copy import copy
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from services.config_loader import load_schema, schema_to_mapping
from services.excel_service import make_json_bytes, make_template_excel_bytes, read_excel_or_csv
from services.oracle_client import OracleFusionClient
from services.payload_builder import PayloadBuildError, build_payload_from_row
from services.validation_service import validation_summary

st.set_page_config(page_title="Upload Runner", page_icon="🚀", layout="wide")

st.title("🚀 Upload Runner")
st.caption("Validasi Excel, preview JSON payload, lalu upload ke Oracle Fusion REST API. Gunakan field ID untuk POST, bukan field Name hasil GET.")

schema = load_schema("inventory_organizations.json")
default_mapping = schema_to_mapping(schema)

st.sidebar.header("Oracle Connection")
try:
    default_base_url = st.secrets.get("ORACLE_BASE_URL", "")
    default_username = st.secrets.get("ORACLE_USERNAME", "")
    default_password = st.secrets.get("ORACLE_PASSWORD", "")
except Exception:
    default_base_url = ""
    default_username = ""
    default_password = ""

base_url = st.sidebar.text_input(
    "Oracle Base URL",
    value=default_base_url,
    placeholder="https://your-instance.oraclecloud.com",
)
username = st.sidebar.text_input("Username", value=default_username)
password = st.sidebar.text_input("Password", value=default_password, type="password")
timeout = st.sidebar.number_input("Timeout per request (seconds)", min_value=10, max_value=300, value=60, step=10)

st.sidebar.header("Run Options")
dry_run = st.sidebar.toggle("Mock / Dry Run mode", value=True, help="Jika aktif, aplikasi tidak mengirim POST ke Oracle.")
upsert_mode = st.sidebar.toggle("Use Upsert-Mode", value=False, help="Jika aktif, header Upsert-Mode:true akan dikirim saat Live mode.")
stop_on_error = st.sidebar.toggle("Stop on first error", value=False)
strict_template = st.sidebar.toggle(
    "Strict template validation",
    value=True,
    help="Jika aktif, urutan dan nama kolom Excel harus sama persis dengan template/mapping.",
)

st.subheader("1. Mapping dan file contoh")
with st.expander("Upload mapping JSON opsional", expanded=False):
    mapping_file = st.file_uploader("Upload mapping JSON dari Template Builder", type=["json"], key="mapping_json")
    if mapping_file:
        mapping = json.load(mapping_file)
        st.success("Mapping JSON berhasil dibaca.")
    else:
        mapping = default_mapping
        st.info("Belum ada mapping JSON yang diupload. Aplikasi memakai default mapping dari schema.")

col_a, col_b = st.columns([1, 2])
with col_a:
    st.write("API aktif")
    st.code(f"{mapping.get('method')} {mapping.get('endpoint')}", language="http")
    if base_url:
        st.write("Target instance")
        st.code(base_url.rstrip("/"), language="text")

with col_b:
    st.write("Field yang dipakai oleh mapping aktif")
    st.dataframe(pd.DataFrame(mapping.get("fields", [])), use_container_width=True, height=220)

st.write("Download file berdasarkan mapping aktif:")
dl1, dl2, dl3 = st.columns(3)
with dl1:
    st.download_button(
        "⬇️ Download Contoh Upload Excel",
        data=make_template_excel_bytes(mapping),
        file_name="inventory_organizations_upload_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with dl2:
    st.download_button(
        "⬇️ Download Mapping JSON",
        data=make_json_bytes(mapping),
        file_name="inventory_organizations_mapping.json",
        mime="application/json",
    )
with dl3:
    try:
        from services.excel_service import fields_to_template_dataframe

        sample_df = fields_to_template_dataframe(mapping, sample_rows=1)
        sample_payload = build_payload_from_row(sample_df.iloc[0], mapping)
    except Exception:
        sample_payload = {}

    st.download_button(
        "⬇️ Download Sample Payload JSON",
        data=make_json_bytes(sample_payload),
        file_name="inventory_organizations_sample_payload.json",
        mime="application/json",
    )

st.subheader("2. Upload Excel / CSV")
uploaded_file = st.file_uploader("Upload file data", type=["xlsx", "csv"])

if not uploaded_file:
    st.info("Download contoh upload Excel di atas, isi sheet Upload_Template, lalu upload kembali di sini.")
    st.stop()

try:
    df = read_excel_or_csv(uploaded_file)
except Exception as exc:
    st.error(f"File gagal dibaca: {exc}")
    st.info("Pastikan file Excel punya sheet bernama Upload_Template. Kalau pakai CSV, pastikan header kolom sama dengan template.")
    st.stop()

# Normalize header spacing only; do not rename content semantically.
df.columns = [str(col).strip() for col in df.columns]

st.write("Preview data")
st.dataframe(df, use_container_width=True)

summary = validation_summary(df, mapping, strict_template=strict_template)

st.subheader("3. Validasi")
col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Total rows", summary["total_rows"])
col2.metric("Missing columns", len(summary["missing_columns"]))
col3.metric("Extra columns", len(summary["extra_columns"]))
col4.metric("Required errors", summary["required_error_count"])
col5.metric("Read-only errors", summary["read_only_error_count"])
col6.metric("Payload errors", summary["payload_error_count"])

if summary["exact_match"]:
    st.success("Struktur kolom Excel sudah sama persis dengan template/mapping aktif.")
else:
    st.warning("Struktur kolom Excel belum sama dengan template/mapping aktif.")

    c1, c2 = st.columns(2)
    with c1:
        st.write("Expected columns dari template/mapping")
        st.dataframe(pd.DataFrame({"expected_columns": summary["expected_columns"]}), use_container_width=True, height=260)
    with c2:
        st.write("Actual columns dari file yang diupload")
        st.dataframe(pd.DataFrame({"actual_columns": summary["actual_columns"]}), use_container_width=True, height=260)

if summary["missing_columns"]:
    st.error("Kolom kurang:")
    st.write(summary["missing_columns"])

if summary["extra_columns"]:
    st.error("Kolom tambahan yang tidak ada di mapping:")
    st.write(summary["extra_columns"])

if not summary["order_is_same"] and not summary["missing_columns"] and not summary["extra_columns"]:
    st.warning("Nama kolom sudah lengkap, tapi urutannya berbeda dari template.")

if not summary["required_errors"].empty:
    with st.expander("Required value errors", expanded=True):
        st.dataframe(summary["required_errors"], use_container_width=True)

if not summary["read_only_errors"].empty:
    st.error("Ada field read-only yang tidak boleh dikirim saat POST. Ganti ke field ID sesuai saran.")
    with st.expander("Read-only field errors", expanded=True):
        st.dataframe(summary["read_only_errors"], use_container_width=True)

if not summary["payload_errors"].empty:
    with st.expander("Payload build errors", expanded=True):
        st.dataframe(summary["payload_errors"], use_container_width=True)

if not summary["is_valid"]:
    st.warning("Perbaiki error validasi sebelum upload live. Preview payload masih bisa dicoba untuk row tertentu kalau datanya valid.")
else:
    st.success("Data valid untuk dibentuk menjadi JSON payload.")

st.subheader("4. Preview dan download payload JSON")
row_number = st.number_input("Pilih nomor row DataFrame untuk preview/test", min_value=0, max_value=max(len(df) - 1, 0), value=0)
try:
    selected_payload = build_payload_from_row(df.iloc[int(row_number)], mapping)
    st.json(selected_payload)
    st.download_button(
        "⬇️ Download JSON Row Ini",
        data=make_json_bytes(selected_payload),
        file_name=f"payload_row_{int(row_number) + 2}.json",
        mime="application/json",
    )
except PayloadBuildError as exc:
    selected_payload = None
    st.error(f"Row ini belum bisa jadi payload: {exc}")

payloads_for_download: List[Dict[str, Any]] = []
payload_download_errors: List[Dict[str, Any]] = []
for idx, row in df.iterrows():
    try:
        payloads_for_download.append({
            "excel_row": int(idx) + 2,
            "payload": build_payload_from_row(row, mapping),
        })
    except PayloadBuildError as exc:
        payload_download_errors.append({"excel_row": int(idx) + 2, "error": str(exc)})

if payload_download_errors:
    with st.expander("Row yang belum bisa dibuat JSON", expanded=False):
        st.dataframe(pd.DataFrame(payload_download_errors), use_container_width=True)
else:
    st.download_button(
        "⬇️ Download Semua Payload JSON",
        data=make_json_bytes(payloads_for_download),
        file_name="inventory_organizations_all_payloads.json",
        mime="application/json",
    )


def make_client() -> OracleFusionClient:
    return OracleFusionClient(base_url=base_url, username=username, password=password, timeout=int(timeout))


def extract_message(body: Any) -> str:
    if isinstance(body, dict):
        for key in ["title", "detail", "message", "error", "o:errorDetails"]:
            if key in body:
                return str(body[key])
        return json.dumps(body, ensure_ascii=False)[:3000]
    return str(body)[:3000]


def log_to_excel_bytes(log_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        log_df.to_excel(writer, sheet_name="Upload_Log", index=False)
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                font = copy(cell.font)
                font.bold = True
                cell.font = font
            for column_cells in ws.columns:
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 60)
    return output.getvalue()


st.subheader("5. Test dan Run")
if not dry_run and (not base_url or not username or not password):
    st.warning("Untuk Live mode, isi Oracle Base URL, username, dan password di sidebar.")

button_col1, button_col2, button_col3 = st.columns(3)

with button_col1:
    if st.button("🔌 Test Connection", disabled=dry_run or not base_url or not username or not password):
        try:
            response = make_client().test_connection(mapping["endpoint"])
            if response.ok:
                st.success(f"Connection OK. Status {response.status_code}")
            else:
                st.error(f"Connection gagal. Status {response.status_code}")
            st.json(response.body if isinstance(response.body, dict) else {"response": response.body})
        except Exception as exc:
            st.error(f"Connection error: {exc}")

with button_col2:
    test_disabled = selected_payload is None or (not dry_run and (not base_url or not username or not password))
    if st.button("🧪 Test Selected Row", disabled=test_disabled):
        if dry_run:
            st.success("Dry run berhasil. Payload tidak dikirim ke Oracle.")
            st.json(selected_payload)
        else:
            try:
                response = make_client().post(mapping["endpoint"], selected_payload, upsert_mode=upsert_mode)
                if response.ok:
                    st.success(f"POST berhasil. Status {response.status_code}")
                else:
                    st.error(f"POST gagal. Status {response.status_code}")
                st.json(response.body if isinstance(response.body, dict) else {"response": response.body})
            except Exception as exc:
                st.error(f"POST error: {exc}")

with button_col3:
    run_disabled = not summary["is_valid"] or (not dry_run and (not base_url or not username or not password))
    run_all = st.button("🚀 Run All Rows", disabled=run_disabled)

if run_all:
    logs: List[Dict[str, Any]] = []
    progress = st.progress(0)
    status_placeholder = st.empty()
    client = make_client() if not dry_run else None

    for idx, row in df.iterrows():
        excel_row = int(idx) + 2
        status_placeholder.write(f"Processing Excel row {excel_row}...")
        try:
            payload = build_payload_from_row(row, mapping)
            display_key = payload.get(mapping.get("display_key_field", ""), "")

            if dry_run:
                logs.append({
                    "excel_row": excel_row,
                    "display_key": display_key,
                    "status_code": "DRY_RUN",
                    "success": True,
                    "response_id": "",
                    "message": "Dry run: payload valid, not posted to Oracle.",
                    "request_payload": json.dumps(payload, ensure_ascii=False),
                    "response_body": "",
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                })
            else:
                response = client.post(mapping["endpoint"], payload, upsert_mode=upsert_mode)
                body = response.body
                response_id = body.get(mapping.get("response_id_field")) if isinstance(body, dict) else ""
                logs.append({
                    "excel_row": excel_row,
                    "display_key": display_key,
                    "status_code": response.status_code,
                    "success": response.ok,
                    "response_id": response_id,
                    "message": extract_message(body),
                    "request_payload": json.dumps(payload, ensure_ascii=False),
                    "response_body": json.dumps(body, ensure_ascii=False) if isinstance(body, dict) else str(body),
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                })
                if stop_on_error and not response.ok:
                    break
        except Exception as exc:
            logs.append({
                "excel_row": excel_row,
                "display_key": "",
                "status_code": "ERROR",
                "success": False,
                "response_id": "",
                "message": str(exc),
                "request_payload": "",
                "response_body": "",
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            if stop_on_error:
                break
        progress.progress((idx + 1) / len(df))

    status_placeholder.write("Selesai.")
    log_df = pd.DataFrame(logs)
    st.subheader("Upload Log")
    st.dataframe(log_df, use_container_width=True)

    st.download_button(
        "⬇️ Download Log Excel",
        data=log_to_excel_bytes(log_df),
        file_name=f"inventory_org_upload_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
