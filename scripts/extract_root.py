#!/usr/bin/env python3

from __future__ import annotations

import html
import json
import re
from collections import OrderedDict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile


SOURCE_ODS = Path(
    "/Users/philstilwell/Google Drive/Desktop/oddsNends/photoshopWork/BLOGS/LogFall/Fallacies.ods"
)
OUTPUT_JSON = Path(
    "/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/fallacies.json"
)
EDITORIAL_OVERRIDES_PATH = Path(
    "/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/editorial_overrides.json"
)

ODS_NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

CATEGORY_ORDER = [
    "Formal",
    "Mathematical",
    "Causal",
    "Linguistic",
    "Conceptual",
    "Evidential",
    "Perceptual",
    "Perspectival",
    "Epistemic",
    "Tactical",
    "Emotional",
]
VALID_CATEGORIES = set(CATEGORY_ORDER)

EXCLUDED_NAMES = {
    "Conditional or questionable fallacies",
    "Correlative based fallacies",
    "Formal Fallacies",
    "Informal fallacies",
    "Propositional fallacies",
    "Quantificational fallacies",
    "Selection bias",
    "Syllogistic fallacies",
    "Entitlement conjuring",
    "Fallacy of the transposed conditional.",
}

DUPLICATE_NAME_PATTERN = re.compile(r"\s+\(duplicate\)$", re.IGNORECASE)

EXAMPLE_OVERRIDES = {
    "Broken window fallacy": (
        "After downtown vandalism, a mayor praises the damage for 'creating jobs' for glaziers "
        "while ignoring the shops, wages, and purchases that would have existed if the windows "
        "had never been broken."
    ),
    "Definist fallacy": (
        "Justice is simply whatever the law permits, so once a policy is legal it is automatically just."
    ),
    "False attribution": (
        "A viral post claims a Nobel Prize winner proved a miracle cure works, but the quote is "
        "misattributed and the researcher never made the claim."
    ),
    "Luddite fallacy": (
        "A city blocks warehouse automation on the ground that any labor-saving technology must "
        "reduce total employment, without considering how demand and new roles can shift elsewhere."
    ),
}

CATEGORY_OVERRIDES = {
    "Broken window fallacy": ["Evidential", "Mathematical"],
    "Luddite fallacy": ["Evidential"],
}


def normalize_text(value: str) -> str:
    value = value or ""
    replacements = {
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u00a0": " ",
        "\u2026": "...",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = value.replace("\r", "\n")
    value = re.sub(r"\s+\n", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    value = re.sub(r"[ \t]{2,}", " ", value)
    return value.strip()


def htmlish_to_text(value: str) -> str:
    value = value or ""
    replacements = {
        r"(?i)<br\s*/?>": "\n",
        r"(?i)<hr\s*/?>": "\n\n",
        r"(?i)</p>": "\n\n",
        r"(?i)</blockquote>": "\n\n",
        r"(?i)<li>": "- ",
        r"(?i)</li>": "\n",
    }
    for pattern, replacement in replacements.items():
        value = re.sub(pattern, replacement, value)
    value = re.sub(r"<[^>]+>", "", value)
    value = html.unescape(value)
    value = normalize_text(value)
    value = re.sub(r" *\n *", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def sentence_case(value: str) -> str:
    value = normalize_text(value)
    if not value:
        return value
    if value.startswith("where "):
        value = "Occurs when " + value[6:]
    elif value.startswith("when "):
        value = "Occurs when " + value[5:]
    elif value.startswith("occurs when "):
        value = "Occurs when " + value[12:]
    else:
        value = value[0].upper() + value[1:]
    return ensure_terminal_punctuation(value)


def ensure_terminal_punctuation(value: str) -> str:
    value = normalize_text(value)
    if not value:
        return value
    if re.search(r"[.!?]['\")\]]*$", value):
        return value
    return value + "."


def clean_example(value: str) -> str:
    value = htmlish_to_text(value)
    if not value or value == "The example will go here.":
        return ""
    return ensure_terminal_punctuation(value)


def clean_block(value: str) -> str:
    value = htmlish_to_text(value)
    if not value:
        return ""
    return ensure_terminal_punctuation(value)


def slugify(value: str) -> str:
    value = normalize_text(value).lower()
    value = value.replace("&", " and ")
    value = value.replace("/", " ")
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = re.sub(r"-{2,}", "-", value)
    return value.strip("-")


def read_sheet_rows(path: Path, sheet_name: str) -> list[list[str]]:
    with ZipFile(path) as archive:
        xml = archive.read("content.xml")
    root = ET.fromstring(xml)
    spreadsheet = root.find("office:body", ODS_NS).find("office:spreadsheet", ODS_NS)
    for table in spreadsheet.findall("table:table", ODS_NS):
        if table.attrib.get(f"{{{ODS_NS['table']}}}name") != sheet_name:
            continue
        rows: list[list[str]] = []
        for row in table.findall("table:table-row", ODS_NS):
            row_repeat = int(row.attrib.get(f"{{{ODS_NS['table']}}}number-rows-repeated", "1"))
            cells: list[str] = []
            for cell in row.findall("table:table-cell", ODS_NS):
                col_repeat = int(
                    cell.attrib.get(f"{{{ODS_NS['table']}}}number-columns-repeated", "1")
                )
                paragraphs = ["".join(p.itertext()) for p in cell.findall("text:p", ODS_NS)]
                value = "\n".join(paragraphs)
                cells.extend([value] * col_repeat)
            for _ in range(row_repeat):
                rows.append(cells)
        return [row for row in rows if any(cell.strip() for cell in row)]
    raise RuntimeError(f"{sheet_name} sheet not found in source workbook.")


def load_wordpress_inventory(path: Path) -> set[str]:
    rows = read_sheet_rows(path, "Wordpress LogFall")
    published = set()
    for row in rows[1:]:
        if len(row) <= 11:
            continue
        name = normalize_text(row[5])
        definition = normalize_text(row[10])
        example = normalize_text(row[11])
        if not name or "<a href=" in name or not definition or not example:
            continue
        published.add(name)
    return published


def load_editorial_overrides(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    raw = json.loads(path.read_text())
    return raw.get("records", {})


def collect_categories(*values: str) -> list[str]:
    seen = OrderedDict()
    for value in values:
        value = normalize_text(value)
        if value in VALID_CATEGORIES:
            seen[value] = None
    return list(seen.keys())


def merge_unique(values: Iterable[str]) -> list[str]:
    seen = OrderedDict()
    for value in values:
        value = clean_block(value)
        if value:
            seen[value] = None
    return list(seen.keys())


def longest_text(*values: str) -> str:
    cleaned = [normalize_text(value) for value in values if normalize_text(value)]
    if not cleaned:
        return ""
    return max(cleaned, key=len)


def apply_editorial_override(record: dict, override: dict) -> dict:
    updated = dict(record)
    if "categories" in override:
        updated["categories"] = [category for category in override["categories"] if category in VALID_CATEGORIES]
    if "aliases" in override:
        updated["aliases"] = [normalize_text(alias) for alias in override["aliases"] if normalize_text(alias)]
    for key in ["originalNumber", "family", "subCategory", "subSubCategory", "editorialStatus"]:
        if key in override:
            updated[key] = normalize_text(override[key])
    if "definition" in override:
        updated["definition"] = sentence_case(override["definition"])
    if "example" in override:
        updated["example"] = clean_example(override["example"])
    if "notes" in override:
        updated["notes"] = clean_block(override["notes"])
    if "caseStudies" in override:
        updated["caseStudies"] = merge_unique(override["caseStudies"])
    return updated


def build_records(
    rows: list[list[str]], wordpress_inventory: set[str], editorial_overrides: dict[str, dict]
) -> list[dict]:
    records_by_name: OrderedDict[str, dict] = OrderedDict()
    for row in rows[1:]:
        name = normalize_text(row[9] if len(row) > 9 else "")
        if (
            not name
            or name in EXCLUDED_NAMES
            or DUPLICATE_NAME_PATTERN.search(name)
            or name not in wordpress_inventory
        ):
            continue

        definition = sentence_case(row[11] if len(row) > 11 else "")
        example = clean_example(row[12] if len(row) > 12 else "")
        notes = clean_block(row[13] if len(row) > 13 else "")
        cases = merge_unique(row[index] if len(row) > index else "" for index in range(14, 19))

        if not example and name in EXAMPLE_OVERRIDES:
            example = clean_example(EXAMPLE_OVERRIDES[name])

        if not definition or not example:
            continue

        categories = collect_categories(
            row[2] if len(row) > 2 else "",
            row[3] if len(row) > 3 else "",
            row[4] if len(row) > 4 else "",
        )
        if not categories and name in CATEGORY_OVERRIDES:
            categories = CATEGORY_OVERRIDES[name]

        if not categories:
            continue

        aliases = [
            alias.strip()
            for alias in re.split(r"[,/;]|\s+or\s+", normalize_text(row[10] if len(row) > 10 else ""))
            if alias.strip()
        ]

        record = {
            "name": name.rstrip("."),
            "slug": slugify(name),
            "categories": categories,
            "originalNumber": normalize_text(row[5] if len(row) > 5 else ""),
            "family": normalize_text(row[6] if len(row) > 6 else ""),
            "subCategory": normalize_text(row[7] if len(row) > 7 else ""),
            "subSubCategory": normalize_text(row[8] if len(row) > 8 else ""),
            "aliases": aliases,
            "definition": definition,
            "example": example,
            "notes": notes,
            "caseStudies": cases,
            "editorialStatus": "cleaned-from-root",
        }

        existing = records_by_name.get(record["name"])
        if not existing:
            records_by_name[record["name"]] = record
            continue

        merged_categories = OrderedDict((category, None) for category in existing["categories"])
        for category in record["categories"]:
            merged_categories[category] = None

        existing["categories"] = sorted(
            merged_categories.keys(),
            key=lambda item: CATEGORY_ORDER.index(item) if item in CATEGORY_ORDER else 999,
        )
        existing["aliases"] = sorted(set(existing["aliases"]) | set(record["aliases"]))
        existing["definition"] = sentence_case(longest_text(existing["definition"], record["definition"]))
        existing["example"] = clean_example(longest_text(existing["example"], record["example"]))
        existing["notes"] = clean_block(longest_text(existing["notes"], record["notes"]))
        existing["caseStudies"] = merge_unique(existing["caseStudies"] + record["caseStudies"])

    records = [apply_editorial_override(record, editorial_overrides.get(record["name"], {})) for record in records_by_name.values()]
    records.sort(key=lambda item: item["name"].lower())
    return records


def build_category_summary(records: list[dict]) -> list[dict]:
    summary = []
    for category in CATEGORY_ORDER:
        members = [record["name"] for record in records if category in record["categories"]]
        if not members:
            continue
        summary.append(
            {
                "name": category,
                "slug": slugify(category),
                "count": len(members),
                "members": members,
            }
        )
    return summary


def main() -> None:
    rows = read_sheet_rows(SOURCE_ODS, "ROOT")
    wordpress_inventory = load_wordpress_inventory(SOURCE_ODS)
    editorial_overrides = load_editorial_overrides(EDITORIAL_OVERRIDES_PATH)
    records = build_records(rows, wordpress_inventory, editorial_overrides)
    payload = {
        "source": SOURCE_ODS.name,
        "sheet": "ROOT",
        "editorialOverridesSource": "data/editorial_overrides.json",
        "recordCount": len(records),
        "categories": build_category_summary(records),
        "records": records,
    }
    OUTPUT_JSON.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps({"output": str(OUTPUT_JSON), "recordCount": len(records)}, indent=2))


if __name__ == "__main__":
    main()
