#!/usr/bin/env python3

from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZipFile


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE_ZIP = PROJECT_ROOT / "fallacy_rationality_tool.zip"
OUTPUT_JSON = PROJECT_ROOT / "data" / "rationality_enrichment.json"

FIELDS = [
    "rationalityDanger",
    "dynamicsToNotice",
    "interactiveMechanic",
    "userAction",
    "feedbackLogic",
    "repairPrompts",
    "warningSigns",
]


def load_tool_payload(path: Path) -> dict:
    with ZipFile(path) as archive:
        return json.loads(archive.read("fallacies-data.json"))


def main() -> None:
    payload = load_tool_payload(SOURCE_ZIP)

    records = {}
    for record in payload.get("fallacies", []):
        name = record.get("name", "").strip()
        if not name:
            continue
        records[name] = {field: (record.get(field, "") or "").strip() for field in FIELDS}

    output = {
        "sourceZip": SOURCE_ZIP.name,
        "records": records,
        "categoryProfiles": payload.get("categoryProfiles", {}),
    }

    OUTPUT_JSON.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps({"output": str(OUTPUT_JSON), "recordCount": len(records)}, indent=2))


if __name__ == "__main__":
    main()
