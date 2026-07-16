from __future__ import annotations

import io
import json
from copy import copy
from typing import Any, Dict, List, Tuple

import pandas as pd

from services.oracle_client import OracleFusionClient, OracleResponse
from services.payload_builder import is_blank, split_payload_path

UNIQUE_EXCEL_COLUMNS = {"OrganizationCode", "OrganizationName"}
SYSTEM_RESPONSE_KEYS = {
    "links",
    "CreatedBy",
    "CreationDate",
    "LastUpdatedBy",
    "LastUpdateDate",
    "LastUpdateLogin",
}

# Endpoint disimpan terpusat supaya mudah diganti kalau instance Oracle punya LOV berbeda.
REFERENCE_ENDPOINTS: Dict[str, Dict[str, Any]] = {
    "Business Units": {
        "endpoint": "/fscmRestApi/resources/11.13.18.05/finBusinessUnitsLOV",
        "id_candidates": ["BusinessUnitId", "BUId"],
        "name_candidates": ["Name", "BusinessUnitName", "BusinessUnit"],
    },
    "Legal Entities": {
        "endpoint": "/fscmRestApi/resources/11.13.18.05/legalEntitiesLOV",
        "id_candidates": ["LegalEntityId", "LegalEntityIdentifier", "LEId"],
        "name_candidates": ["Name", "LegalEntityName"],
    },
    "Inventory Organizations LOV": {
        "endpoint": "/fscmRestApi/resources/11.13.18.05/inventoryOrganizationsLOV",
        "id_candidates": ["OrganizationId", "InventoryOrganizationId"],
        "name_candidates": ["OrganizationName", "OrganizationCode", "Name"],
    },
    # Endpoint lokasi bisa berbeda antar module/setup. Kalau gagal, app tetap jalan dan tampilkan errornya.
    "Locations LOV": {
        "endpoint": "/hcmRestApi/resources/11.13.18.05/locationsLov",
        "id_candidates": ["LocationId", "LocationID"],
        "name_candidates": ["LocationName", "Name", "LocationCode", "AddressLine1"],
        "optional": True,
    },
}


def _as_child_list(value: Any) -> List[Any]:
    """Oracle child resources can appear as a list or as {'items': [...]} after expand."""
    if isinstance(value, list):
        return value
    if isinstance(value, dict) and isinstance(value.get("items"), list):
        return value["items"]
    return []


def get_value_by_payload_path(item: Dict[str, Any], payload_path: str) -> Any:
    """Read a value from an Oracle response item using our mapping payload_path."""
    current: Any = item

    for name, index in split_payload_path(payload_path):
        if current is None:
            return None

        if index is None:
            if not isinstance(current, dict):
                return None
            current = current.get(name)
            continue

        if not isinstance(current, dict):
            return None

        child_list = _as_child_list(current.get(name))
        if index >= len(child_list):
            return None
        current = child_list[index]

    return current


def flatten_item_to_excel_row(item: Dict[str, Any], mapping: Dict[str, Any]) -> Dict[str, Any]:
    """Convert one Oracle GET response item into one Excel-template row based on mapping."""
    row: Dict[str, Any] = {}

    for field in mapping.get("fields", []):
        excel_col = field.get("excel_column")
        payload_path = field.get("payload_path")
        if not excel_col or not payload_path:
            continue

        value = get_value_by_payload_path(item, payload_path)
        row[excel_col] = "" if is_blank(value) else value

    return row


def mapping_with_reference_defaults(
    mapping: Dict[str, Any],
    reference_row: Dict[str, Any],
    use_reference_defaults: bool = True,
    blank_unique_fields: bool = True,
) -> Dict[str, Any]:
    """Copy mapping and optionally set default values from selected reference data."""
    new_mapping = json.loads(json.dumps(mapping, ensure_ascii=False))

    for field in new_mapping.get("fields", []):
        excel_col = field.get("excel_column")
        if not excel_col:
            continue

        if blank_unique_fields and excel_col in UNIQUE_EXCEL_COLUMNS:
            field.pop("default", None)
            continue

        ref_value = reference_row.get(excel_col)
        if use_reference_defaults and not is_blank(ref_value):
            field["default"] = ref_value

    new_mapping["reference_source"] = {
        "OrganizationId": reference_row.get("OrganizationId", ""),
        "OrganizationCode": reference_row.get("OrganizationCode", ""),
        "OrganizationName": reference_row.get("OrganizationName", ""),
    }
    return new_mapping


def reference_template_dataframe(
    mapping: Dict[str, Any],
    reference_row: Dict[str, Any],
    sample_rows: int = 3,
    blank_unique_fields: bool = True,
) -> pd.DataFrame:
    """Create an Excel template dataframe where the first row is filled from selected reference."""
    data: Dict[str, List[Any]] = {}

    for field in mapping.get("fields", []):
        col = field.get("excel_column")
        if not col:
            continue
        data[col] = ["" for _ in range(sample_rows)]

    df = pd.DataFrame(data).astype(object)
    if df.empty or sample_rows <= 0:
        return df

    for col in df.columns:
        if blank_unique_fields and col == "OrganizationCode":
            df.loc[0, col] = "NEW_ORG_CODE"
        elif blank_unique_fields and col == "OrganizationName":
            df.loc[0, col] = "New Inventory Organization Name"
        else:
            df.loc[0, col] = reference_row.get(col, "")

    return df


def item_preview_dataframe(items: List[Dict[str, Any]]) -> pd.DataFrame:
    """Make GET results readable in Streamlit without dumping links/nested objects into wide table cells."""
    rows: List[Dict[str, Any]] = []

    for idx, item in enumerate(items):
        row: Dict[str, Any] = {"result_index": idx}
        for key, value in item.items():
            if key in SYSTEM_RESPONSE_KEYS:
                continue
            if isinstance(value, (dict, list)):
                row[key] = json.dumps(value, ensure_ascii=False)[:500]
            else:
                row[key] = value
        rows.append(row)

    return pd.DataFrame(rows)


def available_fields_from_reference(mapping: Dict[str, Any], reference_row: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Rows for UI editor: field list + value pulled from selected Oracle reference."""
    rows: List[Dict[str, Any]] = []

    for field in mapping.get("fields", []):
        col = field.get("excel_column")
        ref_value = reference_row.get(col, "")
        has_reference_value = not is_blank(ref_value)
        rows.append({
            "include": bool(field.get("required") or has_reference_value),
            "required": bool(field.get("required")),
            "excel_column": col,
            "payload_path": field.get("payload_path"),
            "type": field.get("type"),
            "reference_value": ref_value,
            "description": field.get("description", ""),
        })

    return rows


def _first_existing_value(item: Dict[str, Any], keys: List[str]) -> Any:
    for key in keys:
        value = item.get(key)
        if not is_blank(value):
            return value
    return ""


def items_to_reference_dataframe(reference_type: str, items: List[Dict[str, Any]], config: Dict[str, Any]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for item in items:
        row: Dict[str, Any] = {
            "ReferenceType": reference_type,
            "ID": _first_existing_value(item, config.get("id_candidates", [])),
            "Name": _first_existing_value(item, config.get("name_candidates", [])),
        }
        for key, value in item.items():
            if key in SYSTEM_RESPONSE_KEYS:
                continue
            if isinstance(value, (dict, list)):
                row[key] = json.dumps(value, ensure_ascii=False)[:800]
            else:
                row[key] = value
        rows.append(row)
    return pd.DataFrame(rows)


def fetch_reference_collection(
    client: OracleFusionClient,
    endpoint: str,
    limit: int = 100,
    q_filter: str = "",
) -> OracleResponse:
    params: Dict[str, Any] = {"limit": int(limit), "totalResults": "true"}
    if q_filter.strip():
        params["q"] = q_filter.strip()
    return client.get_collection_items(endpoint, params=params)


def reference_workbook_bytes(reference_tables: Dict[str, pd.DataFrame], errors_df: pd.DataFrame | None = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wrote_sheet = False
        for sheet_name, df in reference_tables.items():
            safe_sheet = sheet_name[:31]
            if df is None or df.empty:
                continue
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            wrote_sheet = True

        if errors_df is not None and not errors_df.empty:
            errors_df.to_excel(writer, sheet_name="Fetch_Errors", index=False)
            wrote_sheet = True

        if not wrote_sheet:
            pd.DataFrame([{"message": "Tidak ada reference data yang berhasil diambil."}]).to_excel(
                writer, sheet_name="README", index=False
            )

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
