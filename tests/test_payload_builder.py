from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from services.config_loader import MINIMAL_CREATE_IO_COLUMNS, load_schema, schema_to_mapping
from services.payload_builder import build_payload_from_row


def main() -> None:
    schema = load_schema("inventory_organizations.json")
    mapping = schema_to_mapping(schema, MINIMAL_CREATE_IO_COLUMNS)

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
        "ItemGroupingCode": "ORA_RCS_IGB_DFTN",
        # intentionally blank: DFTN should auto-fill this from OrganizationCode
        "ItemDefinitionOrganizationCode": "",
        "invOrgParameters.ScheduleId": 100000016383001,
    })

    payload = build_payload_from_row(row, mapping)

    assert payload["OrganizationCode"] == "INV_ORG_TEST"
    assert payload["InventoryFlag"] is True
    assert payload["ManagementBusinessUnitId"] == 204
    assert payload["LegalEntityId"] == 204
    assert payload["ProfitCenterBusinessUnitId"] == 204
    assert payload["MasterOrganizationId"] == 204
    assert payload["ItemGroupingCode"] == "ORA_RCS_IGB_DFTN"
    assert payload["ItemDefinitionOrganizationCode"] == "INV_ORG_TEST"
    assert payload["invOrgParameters"][0]["ScheduleId"] == 100000016383001

    print("Payload builder OK")
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
