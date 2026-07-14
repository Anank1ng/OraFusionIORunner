from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

from services.config_loader import all_request_fields, load_schema, schema_to_mapping
from services.excel_service import make_json_bytes, make_template_excel_bytes, fields_to_template_dataframe, dictionary_dataframe
from services.payload_builder import build_payload_from_row, PayloadBuildError

st.set_page_config(page_title="Template & JSON Builder", page_icon="🧱", layout="wide")

st.title("🧱 Template & JSON Builder")
st.caption("Pilih field Oracle Fusion, lalu generate template Excel, mapping JSON, dan sample request payload.")

schema = load_schema("inventory_organizations.json")
fields = all_request_fields(schema)

with st.expander("Informasi API", expanded=True):
    col1, col2, col3 = st.columns(3)
    col1.metric("API", schema.get("api_name", "-"))
    col2.metric("Method", schema.get("method", "-"))
    col3.metric("Version", schema.get("version", "-"))
    st.code(schema.get("endpoint", ""), language="http")
    for note in schema.get("notes", []):
        st.write(f"- {note}")

st.subheader("1. Pilih field yang masuk template")

field_rows = []
for f in fields:
    field_rows.append({
        "include": bool(f.get("include_by_default")),
        "required": bool(f.get("required")),
        "excel_column": f.get("excel_column"),
        "payload_path": f.get("payload_path"),
        "type": f.get("type"),
        "default": f.get("default", ""),
        "max_length": f.get("max_length", ""),
        "description": f.get("description", "")
    })

field_df = pd.DataFrame(field_rows)
edited_df = st.data_editor(
    field_df,
    hide_index=True,
    use_container_width=True,
    disabled=["required", "excel_column", "payload_path", "type", "max_length", "description"],
    column_config={
        "include": st.column_config.CheckboxColumn("Include", help="Masukkan field ini ke template Excel"),
        "required": st.column_config.CheckboxColumn("Required"),
        "description": st.column_config.TextColumn("Description", width="large")
    }
)

selected_columns = edited_df.loc[edited_df["include"] == True, "excel_column"].tolist()
mapping = schema_to_mapping(schema, selected_columns)

# Apply editable default values from the data editor.
default_by_col = dict(zip(edited_df["excel_column"], edited_df["default"]))
for field in mapping["fields"]:
    default_value = default_by_col.get(field["excel_column"], field.get("default"))
    if default_value == "":
        field.pop("default", None)
    else:
        field["default"] = default_value

st.subheader("2. Preview template dan payload")
col_left, col_right = st.columns([1.2, 1])

with col_left:
    template_df = fields_to_template_dataframe(mapping, sample_rows=3)
    st.write("Preview template Excel")
    st.dataframe(template_df, use_container_width=True)

with col_right:
    try:
        sample_payload = build_payload_from_row(template_df.iloc[0], mapping)
        st.write("Sample JSON payload dari baris pertama")
        st.json(sample_payload)
    except PayloadBuildError as exc:
        sample_payload = {}
        st.error(f"Sample payload belum valid: {exc}")

st.subheader("3. Download hasil builder")
col1, col2, col3 = st.columns(3)

with col1:
    st.download_button(
        "⬇️ Download Template Excel",
        data=make_template_excel_bytes(mapping),
        file_name="inventory_organizations_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col2:
    st.download_button(
        "⬇️ Download Mapping JSON",
        data=make_json_bytes(mapping),
        file_name="inventory_organizations_mapping.json",
        mime="application/json"
    )

with col3:
    st.download_button(
        "⬇️ Download Sample Payload",
        data=make_json_bytes(sample_payload),
        file_name="inventory_organizations_sample_payload.json",
        mime="application/json"
    )

with st.expander("Field Dictionary"):
    st.dataframe(dictionary_dataframe(mapping), use_container_width=True)
