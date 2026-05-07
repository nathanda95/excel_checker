import json
import re
from pathlib import Path

from excel_checker.utils.text_utils import normalize_key, normalize_text_for_compare
from excel_checker.utils.text_utils import (
    normalize_key,
    normalize_text_for_compare,
    parse_header_metadata,
    clean_header_name,
)


def load_postal_code_rules() -> dict:
    project_root = Path(__file__).resolve().parents[2]
    json_path = project_root / "rule_set.json"

    if not json_path.exists():
        return {}
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not data or not isinstance(data, list):
        return {}

    rows = []

    if isinstance(data, list):
        for block in data:
            if isinstance(block, dict) and "code_postaux" in block:
                rows = block["code_postaux"]
                break

    rules = {}

    for item in rows:
        country = normalize_key(item.get("country", ""))
        country_code = normalize_key(item.get("country_code", ""))
        regex = item.get("regex", "")
        example = item.get("example", "")
        label = item.get("country", "") or item.get("country_code", "")

        if not regex:
            continue

        entry = {
            "country": item.get("country", ""),
            "country_code": item.get("country_code", ""),
            "regex": regex,
            "example": example,
            "label": label,
        }

        if country:
            rules[country] = entry
        if country_code:
            rules[country_code] = entry

    return rules

def _header_label_only(header_value: str) -> str:
    meta = parse_header_metadata(header_value)
    field_name = meta["field_name"] or clean_header_name(header_value)
    return normalize_key(field_name)

def is_postal_code_header(header_value: str) -> bool:
    key = _header_label_only(header_value)
    keywords = [
        "code postal",
        "postal code",
        "zip code",
        "zipcode",
        "postcode",
        "post code",
        "cp",
        "Code postal",
        "Code postal*"
    ]
    return any(word in key for word in keywords)


def is_country_header(header_value: str) -> bool:
    key = _header_label_only(header_value)
    keywords = [
        "pays",
        "country",
        "country region",
        "countryregion",
        "code pays",
        "country code",
        "ship country",
        "bill country",
        "delivery country",
        "Pays/Région*",
        "Pays/Région",
    ]
    return any(word in key for word in keywords)


def normalize_country_lookup(value) -> str:
    return normalize_key(normalize_text_for_compare(value))


def validate_postal_code_for_country(postal_code, country_value, postal_rules: dict):
    postal_text = normalize_text_for_compare(postal_code)
    country_key = normalize_country_lookup(country_value)

    if not postal_text or not country_key:
        return True, None

    rule = postal_rules.get(country_key)
    if not rule:
        return True, None

    try:
        ok = re.fullmatch(rule["regex"], postal_text, flags=re.IGNORECASE) is not None
    except re.error:
        return True, rule

    return ok, rule