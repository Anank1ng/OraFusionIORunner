from __future__ import annotations

import io
import json
from copy import copy
from typing import Any, Dict, List

import pandas as pd


def fields_to_template_dataframe(mapping: Dict[str, Any], sample_rows: int = 3) -> pd.DataFrame:
    data: Dict[str, List[Any]] = {}

    for field in mapping.get("fields", []):
        col = field.get("excel_column")
        if not col:
            continue

        # Template rows after the first one should stay empty.
        # Defaults are applied when building payload, so empty Excel cells can still use mapping defaults.
        data[col] = ["" for _ in range(sample_rows)]

    df = pd.DataFrame(data).astype(object)

    if sample_rows > 0 and not df.empty:
        sample = {
            "OrganizationCode": "INV_ORG_TEST",
            "OrganizationName": "Inventory Org Test",
            "ManagementBusinessUnitId": 204,
            "LegalEntityId": 204,
            "ProfitCenterBusinessUnitId": 204,
            "Status": "Active",
            "LocationId": 1001,
            "InventoryFlag": True,
            "MasterOrganizationId": 204,
            "ManufacturingPlantFlag": True,
            "ContractManufacturingFlag": False,
            "MaintenanceEnabledFlag": False,
            "ItemGroupingCode": "ORA_RCS_IGB_RFRC",
            "ItemDefinitionOrganizationId": 204,
            "invOrgParameters.ScheduleId": 100000016383001,
            "plantParameters.ManufacturingCalendarId": 100000016383001,
            "plantParameters.DefaultSupplySubinventory": "SUB1",
            "plantParameters.DefaultCompletionSubinventory": "SUB1",
            "plantParameters.EnableProcessManufacturingFlag": True,
            "plantParameters.DefaultWorkMethod": "PROCESS_MANUFACTURING",
        }

        defaults = {field.get("excel_column"): field.get("default") for field in mapping.get("fields", [])}

        for col in df.columns:
            if col in sample:
                df.loc[0, col] = sample[col]
            elif defaults.get(col) is not None:
                df.loc[0, col] = defaults[col]

    return df


def dictionary_dataframe(mapping: Dict[str, Any]) -> pd.DataFrame:
    rows = []

    for field in mapping.get("fields", []):
        rows.append({
            "excel_column": field.get("excel_column"),
            "payload_path": field.get("payload_path"),
            "type": field.get("type"),
            "required": field.get("required"),
            "default": field.get("default"),
            "max_length": field.get("max_length"),
            "allowed_values": ", ".join(field.get("allowed_values") or []),
            "description": field.get("description", ""),
        })

    return pd.DataFrame(rows)


def make_template_excel_bytes(mapping: Dict[str, Any], template_df: pd.DataFrame | None = None) -> bytes:
    if template_df is None:
        template_df = fields_to_template_dataframe(mapping, sample_rows=3)
    dict_df = dictionary_dataframe(mapping)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template_df.to_excel(writer, sheet_name="Upload_Template", index=False)
        dict_df.to_excel(writer, sheet_name="Field_Dictionary", index=False)
        pd.DataFrame([
            {"key": "api_name", "value": mapping.get("api_name")},
            {"key": "method", "value": mapping.get("method")},
            {"key": "endpoint", "value": mapping.get("endpoint")},
            {"key": "media_type", "value": mapping.get("supported_media_type")},
            {"key": "note", "value": "Isi sheet Upload_Template lalu upload di halaman Upload Runner."},
        ]).to_excel(writer, sheet_name="README", index=False)

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


def make_json_bytes(data: Dict[str, Any] | List[Any]) -> bytes:
    return json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8")


def read_excel_or_csv(uploaded_file) -> pd.DataFrame:
    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    try:
        return pd.read_excel(uploaded_file, sheet_name="Upload_Template")
    except ValueError:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
