from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

ROOT_DIR = Path(__file__).resolve().parents[1]
SCHEMA_DIR = ROOT_DIR / "schemas"


def list_schema_files() -> List[Path]:
    return sorted(SCHEMA_DIR.glob("*.json"))


def load_schema(schema_name: str = "inventory_organizations.json") -> Dict[str, Any]:
    schema_path = SCHEMA_DIR / schema_name
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema tidak ditemukan: {schema_path}")
    with schema_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def default_fields(schema: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [
        field for field in schema.get("fields", [])
        if field.get("include_by_default") and not field.get("system_generated")
    ]


def all_request_fields(schema: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [field for field in schema.get("fields", []) if not field.get("system_generated")]


def schema_to_mapping(schema: Dict[str, Any], selected_excel_columns: List[str] | None = None) -> Dict[str, Any]:
    selected = set(selected_excel_columns or [f["excel_column"] for f in default_fields(schema)])
    fields = []
    for field in schema.get("fields", []):
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
