from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

ROOT_DIR = Path(__file__).resolve().parents[1]
SCHEMA_DIR = ROOT_DIR / "schemas"

# Minimal fields confirmed from Oracle Fusion UI required (*) fields for creating an Inventory Organization.
# Keep this order so generated Excel templates are easier to read and test.
MINIMAL_CREATE_IO_COLUMNS = [
    "OrganizationCode",
    "OrganizationName",
    "ManagementBusinessUnitId",
    "LegalEntityId",
    "ProfitCenterBusinessUnitId",
    "Status",
    "LocationId",
    "InventoryFlag",
    "MasterOrganizationId",
    "ItemGroupingCode",
    "ItemDefinitionOrganizationCode",
    "invOrgParameters.ScheduleId",
]

STANDARD_WAREHOUSE_SECTIONS = [
    "Core Organization",
    "Financial IDs",
    "Item Definition Settings",
    "Additional Usages",
    "Inventory Settings",
    "Movement Request",
    "Picking Defaults",
]

FULL_ADVANCED_SECTIONS = [
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
]


def list_schema_files() -> List[Path]:
    return sorted(SCHEMA_DIR.glob("*.json"))


def load_schema(schema_name: str = "inventory_organizations.json") -> Dict[str, Any]:
    schema_path = SCHEMA_DIR / schema_name
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema tidak ditemukan: {schema_path}")
    with schema_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def default_fields(schema: Dict[str, Any]) -> List[Dict[str, Any]]:
    fields_by_column = {
        field.get("excel_column"): field
        for field in schema.get("fields", [])
        if field.get("include_by_default") and not field.get("system_generated")
    }
    ordered = [fields_by_column[col] for col in MINIMAL_CREATE_IO_COLUMNS if col in fields_by_column]
    ordered += [field for field in fields_by_column.values() if field not in ordered]
    return ordered


def all_request_fields(schema: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [field for field in schema.get("fields", []) if not field.get("system_generated")]


def schema_to_mapping(schema: Dict[str, Any], selected_excel_columns: List[str] | None = None) -> Dict[str, Any]:
    selected_list = selected_excel_columns or [f["excel_column"] for f in default_fields(schema)]
    selected = set(selected_list)
    field_by_column = {
        field.get("excel_column"): field
        for field in schema.get("fields", [])
        if not field.get("system_generated") and field.get("excel_column")
    }

    # Preserve selected_excel_columns order for generated templates.
    # This keeps presets such as Minimal Create IO predictable.
    ordered_schema_fields = [field_by_column[col] for col in selected_list if col in field_by_column]
    ordered_schema_fields += [
        field for field in schema.get("fields", [])
        if not field.get("system_generated")
        and field.get("excel_column") in selected
        and field not in ordered_schema_fields
    ]

    fields = []
    for field in ordered_schema_fields:
        if field.get("system_generated"):
            continue
        if field.get("excel_column") in selected:
            fields.append({
                "payload_path": field["payload_path"],
                "excel_column": field["excel_column"],
                "type": field.get("type", "string"),
                "required": bool(field.get("required")),
                "default": field.get("default", None),
                "max_length": field.get("max_length"),
                "allowed_values": field.get("allowed_values"),
                "label": field.get("label", field.get("excel_column")),
                "section": field.get("section", "General"),
                "description": field.get("description", ""),
                "reference_hint": field.get("reference_hint", ""),
                "read_only_for_post": bool(field.get("read_only_for_post", False)),
                "use_instead": field.get("use_instead", ""),
            })
    return {
        "api_key": schema.get("api_key"),
        "api_name": schema.get("api_name"),
        "method": schema.get("method", "POST"),
        "endpoint": schema.get("endpoint"),
        "supported_media_type": schema.get("supported_media_type", "application/json"),
        "response_id_field": schema.get("response_id_field", "id"),
        "display_key_field": schema.get("display_key_field"),
        "fields": fields
    }
