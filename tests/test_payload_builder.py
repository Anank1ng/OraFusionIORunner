from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from services.config_loader import load_schema, schema_to_mapping
from services.payload_builder import build_payload_from_row


def main() -> None:
    schema = load_schema("inventory_organizations.json")
    mapping = schema_to_mapping(schema)

    row = pd.Series({
        "OrganizationCode": "INV_ORG_TEST",
        "OrganizationName": "Inventory Org Test",
        "ManagementBusinessUnitId": 204,
        "LegalEntityId": 204,
        "ProfitCenterBusinessUnitId": 204,
        "Status": "Active",
        "LocationId": 1001,
        "InventoryFlag": "true",
        "MasterOrganizationId": 204,
        "ManufacturingPlantFlag": "yes",
        "ContractManufacturingFlag": "no",
        "MaintenanceEnabledFlag": False,
        "ItemGroupingCode": "ORA_RCS_IGB_RFRC",
        "ItemDefinitionOrganizationId": 204,
        "invOrgParameters.ScheduleId": 100000016383001,
        "plantParameters.ManufacturingCalendarId": 100000016383001,
        "plantParameters.DefaultSupplySubinventory": "SUB1",
        "plantParameters.DefaultCompletionSubinventory": "SUB1",
        "plantParameters.EnableProcessManufacturingFlag": "1",
        "plantParameters.DefaultWorkMethod": "PROCESS_MANUFACTURING"
    })

    payload = build_payload_from_row(row, mapping)

    assert payload["OrganizationCode"] == "INV_ORG_TEST"
    assert payload["InventoryFlag"] is True
    assert payload["ManufacturingPlantFlag"] is True
    assert payload["ContractManufacturingFlag"] is False
    assert payload["invOrgParameters"][0]["ScheduleId"] == 100000016383001
    assert payload["plantParameters"][0]["DefaultWorkMethod"] == "PROCESS_MANUFACTURING"

    print("Payload builder OK")
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
