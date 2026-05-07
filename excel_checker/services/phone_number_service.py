import json
import re
from pathlib import Path

from excel_checker.utils.text_utils import (
    normalize_key,
    normalize_text_for_compare,
    parse_header_metadata,
    clean_header_name,
)


def _header_label_only(header_value: str) -> str:
    meta = parse_header_metadata(header_value)
    field_name = meta["field_name"] or clean_header_name(header_value)
    return normalize_key(field_name)


def load_phone_number_rules() -> dict:
    project_root = Path(__file__).resolve().parents[2]
    json_path = project_root / "rule_set.json"

    if not json_path.exists():
        return {}

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    rows = []

    if isinstance(data, list):
        for block in data:
            if isinstance(block, dict) and "phone_number" in block:
                rows = block["phone_number"]
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
            "country_calling_code": item.get("country_calling_code", ""),
            "regex": regex,
            "example": example,
            "label": label,
        }

        if country:
            rules[country] = entry
        if country_code:
            rules[country_code] = entry

    return rules


def is_phone_header(header_value: str) -> bool:
    key = _header_label_only(header_value)

    keywords = [
        "telephone",
        "téléphone",
        "Téléphone",
        "Téléphone*",
        "telephone portable",
        "téléphone portable",
        "telephone mobile",
        "téléphone mobile",
        "mobile",
        "phone",
        "phone number",
        "mobile phone",
        "telephone number",
        "numero de telephone",
        "numéro de téléphone",
        "gsm",
        "tel",
    ]

    return any(word in key for word in keywords)


def normalize_country_lookup(value) -> str:
    key = normalize_key(normalize_text_for_compare(value))

    aliases = {
        "uk": "gb",
        "great britain": "gb",
        "united kingdom": "gb",
        "royaume uni": "gb",

        "france": "fr",
        "espagne": "es",
        "spain": "es",
        "allemagne": "de",
        "germany": "de",
        "italie": "it",
        "italy": "it",
        "portugal": "pt",
        "pologne": "pl",
        "poland": "pl",
        "suede": "se",
        "sweden": "se",
        "suisse": "ch",
        "switzerland": "ch",
        "autriche": "at",
        "austria": "at",
        "belgique": "be",
        "belgium": "be",
        "irlande": "ie",
        "ireland": "ie",
        "pays bas": "nl",
        "netherlands": "nl",
        "republique tcheque": "cz",
        "czech republic": "cz",
        "etats unis": "us",
        "etats unis d amerique": "us",
        "usa": "us",
        "united states": "us",
        "united states of america": "us",
    }

    return aliases.get(key, key)


def normalize_phone_for_check(value) -> str:
    phone = normalize_text_for_compare(value)
    phone = phone.replace(" ", "")
    phone = phone.replace("-", "")
    phone = phone.replace(".", "")
    phone = phone.replace("(", "")
    phone = phone.replace(")", "")
    return phone


def validate_phone_for_country(phone_value, country_value, phone_rules: dict):
    phone_text = normalize_phone_for_check(phone_value)
    country_key = normalize_country_lookup(country_value)

    if not phone_text or not country_key:
        return True, None

    rule = phone_rules.get(country_key)
    if not rule:
        return True, None

    try:
        ok = re.fullmatch(rule["regex"], phone_text, flags=re.IGNORECASE) is not None
    except re.error:
        return True, rule

    return ok, rule