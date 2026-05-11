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
SUPPLEMENTAL_FALLACIES_PATH = Path(
    "/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/supplemental_fallacies.json"
)
RATIONALITY_ENRICHMENT_PATH = Path(
    "/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/rationality_enrichment.json"
)
CASE_STUDY_LIBRARY_PATH = Path(
    "/Users/philstilwell/Documents/Codex/2026-05-09/install-ghostscript/logfall-repo/data/case_study_library.json"
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


def lowercase_initial(value: str) -> str:
    value = normalize_text(value)
    if not value:
        return value
    return value[0].lower() + value[1:]


def extract_occurs_when_clause(value: str) -> str:
    value = clean_block(value)
    if not value:
        return ""

    pattern = re.compile(
        r"(?:^|:\s+)occurs when\s+(?P<core>.+?)\s+The module should make the user see how\b",
        re.IGNORECASE,
    )
    match = pattern.search(value)
    if not match:
        return ""

    return sentence_case(match.group("core"))


def make_reader_friendly_rationality_danger(name: str, value: str) -> str:
    value = clean_block(value)
    if not value:
        return ""

    core = extract_occurs_when_clause(value)
    if not core:
        return value

    core = lowercase_initial(core.rstrip("."))
    return ensure_terminal_punctuation(f"{normalize_text(name)} threatens rationality because {core}")


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


def build_supplemental_records(path: Path, editorial_overrides: dict[str, dict]) -> list[dict]:
    if not path.exists():
        return []

    raw = json.loads(path.read_text())
    records = []
    for entry in raw.get("records", []):
        name = normalize_text(entry.get("name", "")).rstrip(".")
        if not name:
            raise ValueError("Supplemental fallacy record is missing a name.")

        record = {
            "name": name,
            "slug": slugify(name),
            "categories": normalize_categories(
                entry.get("categories", []), context=f'supplemental record "{name}"'
            ),
            "originalNumber": normalize_text(entry.get("originalNumber", "")),
            "family": normalize_text(entry.get("family", "Informal Fallacy")),
            "subCategory": normalize_text(entry.get("subCategory", "")),
            "subSubCategory": normalize_text(entry.get("subSubCategory", "")),
            "aliases": parse_aliases(", ".join(entry.get("aliases", []))),
            "definition": sentence_case(entry.get("definition", "")),
            "example": clean_example(entry.get("example", "")),
            "notes": clean_block(entry.get("notes", "")),
            "caseStudies": merge_unique(entry.get("caseStudies", [])),
            "editorialStatus": normalize_text(
                entry.get("editorialStatus", "added-2026-completeness-pass")
            ),
        }

        if not record["definition"] or not record["example"] or not record["categories"]:
            raise ValueError(f'Supplemental fallacy "{name}" is missing a definition, example, or categories.')

        records.append(apply_editorial_override(record, editorial_overrides.get(name, {})))

    records.sort(key=lambda item: item["name"].lower())
    return records


def merge_record_sets(primary_records: list[dict], supplemental_records: list[dict]) -> list[dict]:
    records_by_name: OrderedDict[str, dict] = OrderedDict(
        (record["name"], record) for record in primary_records
    )
    for record in supplemental_records:
        if record["name"] in records_by_name:
            raise ValueError(f'Duplicate fallacy name across sources: "{record["name"]}".')
        records_by_name[record["name"]] = record
    records = list(records_by_name.values())
    records.sort(key=lambda item: item["name"].lower())
    return records


def collect_categories(*values: str) -> list[str]:
    seen = OrderedDict()
    for value in values:
        value = normalize_text(value)
        if value in VALID_CATEGORIES:
            seen[value] = None
    return normalize_categories(seen.keys())


def normalize_categories(values: Iterable[str], context: str = "record") -> list[str]:
    seen = OrderedDict()
    for value in values:
        value = normalize_text(value)
        if value in VALID_CATEGORIES:
            seen[value] = None
    categories = list(seen.keys())
    if len(categories) > 3:
        raise ValueError(f"{context} has {len(categories)} categories; expected at most 3.")
    return categories


def normalize_category_tags(values: Iterable[str]) -> list[str]:
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


def normalize_manual_case_study(value) -> dict | None:
    if isinstance(value, str):
        summary = clean_block(value)
        if not summary:
            return None
        return {
            "summary": summary,
            "source": "",
            "title": "",
            "date": "",
            "url": "",
        }

    if not isinstance(value, dict):
        return None

    summary = clean_block(value.get("summary", ""))
    if not summary:
        return None

    return {
        "summary": summary,
        "source": normalize_text(value.get("source", "")),
        "title": normalize_text(value.get("title", "")),
        "date": normalize_text(value.get("date", "")),
        "url": normalize_text(value.get("url", "")),
    }


def normalize_manual_case_studies(values: Iterable) -> list[dict]:
    seen = OrderedDict()
    for value in values:
        item = normalize_manual_case_study(value)
        if not item:
            continue
        key = item["url"] or item["summary"]
        seen[key] = item
    return list(seen.values())


def parse_aliases(value: str) -> list[str]:
    seen = OrderedDict()
    for alias in re.split(r"\s*/\s*|\s*;\s*|\s*,\s*", normalize_text(value)):
        alias = alias.strip()
        if alias:
            seen[alias] = None
    return list(seen.keys())


def load_case_study_library(path: Path) -> list[dict]:
    if not path.exists():
        return []

    raw = json.loads(path.read_text())
    entries = []
    for entry in raw.get("entries", []):
        normalized = {
            "id": normalize_text(entry.get("id", "")),
            "summary": clean_block(entry.get("summary", "")),
            "source": normalize_text(entry.get("source", "")),
            "title": normalize_text(entry.get("title", "")),
            "date": normalize_text(entry.get("date", "")),
            "url": normalize_text(entry.get("url", "")),
            "categories": normalize_category_tags(entry.get("categories", [])),
            "families": [normalize_text(value) for value in entry.get("families", []) if normalize_text(value)],
            "subCategories": [normalize_text(value) for value in entry.get("subCategories", []) if normalize_text(value)],
            "subSubCategories": [normalize_text(value) for value in entry.get("subSubCategories", []) if normalize_text(value)],
            "fallacies": [normalize_text(value) for value in entry.get("fallacies", []) if normalize_text(value)],
        }

        if not all([normalized["id"], normalized["summary"], normalized["source"], normalized["title"], normalized["url"]]):
            raise ValueError(f"Case study library entry is missing required fields: {entry}")

        entries.append(normalized)

    return entries


def load_rationality_enrichment(path: Path) -> dict:
    if not path.exists():
        return {"records": {}, "categoryProfiles": {}}

    raw = json.loads(path.read_text())
    records = {}
    for name, values in raw.get("records", {}).items():
        clean_name = normalize_text(name)
        if not clean_name:
            continue
        records[clean_name] = {
            "rationalityDanger": make_reader_friendly_rationality_danger(
                clean_name, values.get("rationalityDanger", "")
            ),
            "mainReasoningProblem": clean_block(
                values.get("mainReasoningProblem", "")
            )
            or extract_occurs_when_clause(values.get("rationalityDanger", "")),
            "dynamicsToNotice": clean_block(values.get("dynamicsToNotice", "")),
            "interactiveMechanic": clean_block(values.get("interactiveMechanic", "")),
            "userAction": clean_block(values.get("userAction", "")),
            "feedbackLogic": clean_block(values.get("feedbackLogic", "")),
            "repairPrompts": clean_block(values.get("repairPrompts", "")),
            "warningSigns": clean_block(values.get("warningSigns", "")),
        }

    category_profiles = {}
    for category, values in raw.get("categoryProfiles", {}).items():
        clean_category = normalize_text(category)
        if clean_category not in VALID_CATEGORIES:
            continue
        category_profiles[clean_category] = {
            "distortion": clean_block(values.get("distortion", "")),
            "danger": clean_block(values.get("danger", "")),
            "mechanic": clean_block(values.get("mechanic", "")),
            "user_action": clean_block(values.get("user_action", "")),
            "feedback": clean_block(values.get("feedback", "")),
        }

    return {"records": records, "categoryProfiles": category_profiles}


def apply_rationality_enrichment(record: dict, enrichment: dict) -> dict:
    updated = dict(record)
    for key in [
        "rationalityDanger",
        "mainReasoningProblem",
        "dynamicsToNotice",
        "interactiveMechanic",
        "userAction",
        "feedbackLogic",
        "repairPrompts",
        "warningSigns",
    ]:
        updated[key] = clean_block(enrichment.get(key, "")) if enrichment else ""
    return updated


def select_case_studies(record: dict, case_study_library: list[dict], limit: int = 5) -> list[dict]:
    manual_cases = normalize_manual_case_studies(record.get("caseStudies", []))
    selected = []
    seen_keys = set()

    def append_case(item: dict) -> None:
        key = item.get("url") or item.get("summary")
        if not key or key in seen_keys or len(selected) >= limit:
            return
        seen_keys.add(key)
        selected.append(item)

    for case in manual_cases:
        if case.get("source") or case.get("url"):
            append_case(case)

    for entry in case_study_library:
        if record["name"] not in entry.get("fallacies", []):
            continue
        append_case(
            {
                "summary": entry["summary"],
                "source": entry["source"],
                "title": entry["title"],
                "date": entry["date"],
                "url": entry["url"],
            }
        )

    for case in manual_cases:
        append_case(case)

    return selected[:limit]


def enrich_case_studies(records: list[dict], case_study_library: list[dict]) -> list[dict]:
    enriched = []
    for record in records:
        updated = dict(record)
        updated["caseStudies"] = select_case_studies(updated, case_study_library)
        enriched.append(updated)
    return enriched


def longest_text(*values: str) -> str:
    cleaned = [normalize_text(value) for value in values if normalize_text(value)]
    if not cleaned:
        return ""
    return max(cleaned, key=len)


def apply_editorial_override(record: dict, override: dict) -> dict:
    updated = dict(record)
    if "categories" in override:
        override_categories = normalize_categories(
            override["categories"], context=f'editorial override for "{updated["name"]}"'
        )
        if override_categories:
            updated["categories"] = override_categories
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
        updated["caseStudies"] = list(override["caseStudies"])
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
            categories = normalize_categories(
                CATEGORY_OVERRIDES[name], context=f'category override for "{name}"'
            )

        if not categories:
            continue

        aliases = parse_aliases(row[10] if len(row) > 10 else "")

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

        existing["categories"] = normalize_categories(
            merged_categories.keys(),
            context=f'merged categories for "{existing["name"]}"',
        )
        merged_aliases = OrderedDict((alias, None) for alias in existing["aliases"])
        for alias in record["aliases"]:
            merged_aliases[alias] = None
        existing["aliases"] = list(merged_aliases.keys())
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
    case_study_library = load_case_study_library(CASE_STUDY_LIBRARY_PATH)
    rationality_enrichment = load_rationality_enrichment(RATIONALITY_ENRICHMENT_PATH)
    records = build_records(rows, wordpress_inventory, editorial_overrides)
    supplemental_records = build_supplemental_records(
        SUPPLEMENTAL_FALLACIES_PATH, editorial_overrides
    )
    records = merge_record_sets(records, supplemental_records)
    records = [
        apply_rationality_enrichment(
            record, rationality_enrichment["records"].get(record["name"], {})
        )
        for record in records
    ]
    records = enrich_case_studies(records, case_study_library)
    payload = {
        "source": SOURCE_ODS.name,
        "sheet": "ROOT",
        "editorialOverridesSource": "data/editorial_overrides.json",
        "supplementalSource": "data/supplemental_fallacies.json",
        "rationalityEnrichmentSource": "data/rationality_enrichment.json",
        "caseStudyLibrarySource": "data/case_study_library.json",
        "recordCount": len(records),
        "categories": build_category_summary(records),
        "categoryProfiles": rationality_enrichment["categoryProfiles"],
        "records": records,
    }
    OUTPUT_JSON.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps({"output": str(OUTPUT_JSON), "recordCount": len(records)}, indent=2))


if __name__ == "__main__":
    main()
