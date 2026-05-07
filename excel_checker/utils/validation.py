import re

from excel_checker.config import (
    EMAIL_REGEX_DEFAULT,
    PHONE_REGEX_DEFAULT,
    ALPHANUM_REGEX_DEFAULT,
)
from excel_checker.utils.text_utils import (
    is_empty,
    normalize_text_for_compare,
    parse_header_metadata,
)


def is_email_header(header_value: str) -> bool:
    meta = parse_header_metadata(header_value)

    field_name = (meta["field_name"] or "").lower()
    field_type = (meta["field_type"] or "").lower()

    email_keywords = ["email", "e-mail", "mail"]

    return (
        "email" in field_type
        or any(keyword in field_name for keyword in email_keywords)
    )


def is_valid_email(value, compiled_regex) -> bool:
    if is_empty(value):
        return True

    email_text = normalize_text_for_compare(value)
    return compiled_regex.fullmatch(email_text) is not None


def normalize_type_name(type_name: str) -> str:
    if not type_name:
        return ""

    text = str(type_name).strip().lower()
    text = text.replace("é", "e").replace("è", "e").replace("ê", "e")
    text = text.replace("à", "a").replace("â", "a")
    text = text.replace("î", "i").replace("ï", "i")
    text = text.replace("ô", "o")
    text = text.replace("ù", "u").replace("û", "u").replace("ü", "u")
    return text


def canonical_type_name(type_name: str) -> str:
    t = normalize_type_name(type_name)

    mapping = {
        "texte": "text",
        "text": "text",
        "string": "text",

        "email": "email",
        "e-mail": "email",
        "mail": "email",

        "numerique": "numeric",
        "numeric": "numeric",
        "number": "numeric",
        "nombre": "numeric",

        "entier": "integer",
        "integer": "integer",
        "int": "integer",

        "decimal": "decimal",
        "float": "decimal",
        "double": "decimal",
        "montant": "decimal",

        "date": "date",
        "datetime": "date",

        "booleen": "boolean",
        "bool": "boolean",
        "boolean": "boolean",
        "oui/non": "boolean",
        "yes/no": "boolean",

        "telephone": "phone",
        "phone": "phone",
        "tel": "phone",
        "mobile": "phone",

        "alphanumerique": "alphanum",
        "alphanumeric": "alphanum",
        "alphanum": "alphanum",
        "code": "alphanum",
    }

    return mapping.get(t, "")


def is_valid_integer(value_text: str) -> bool:
    return re.fullmatch(r"[+-]?\d+", value_text) is not None


def is_valid_decimal(value_text: str) -> bool:
    normalized = value_text.replace(",", ".").strip()
    return re.fullmatch(r"[+-]?\d+(\.\d+)?", normalized) is not None


def is_valid_date(value_text: str) -> bool:
    value_text = value_text.strip()

    patterns = [
        r"\d{2}/\d{2}/\d{4}",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{2}-\d{2}-\d{4}",
        r"\d{2}\.\d{2}\.\d{4}",
    ]

    return any(re.fullmatch(pattern, value_text) for pattern in patterns)


def is_valid_boolean(value_text: str) -> bool:
    valid_values = {
        "true", "false",
        "oui", "non",
        "yes", "no",
        "1", "0",
        "vrai", "faux"
    }
    return value_text.strip().lower() in valid_values


def is_valid_phone(value_text: str) -> bool:
    return re.fullmatch(PHONE_REGEX_DEFAULT, value_text.strip()) is not None


def is_valid_alphanum(value_text: str) -> bool:
    return re.fullmatch(ALPHANUM_REGEX_DEFAULT, value_text.strip()) is not None


def is_type_valid(field_type: str, value, email_regex: str = EMAIL_REGEX_DEFAULT):
    if is_empty(value):
        return True, field_type

    value_text = normalize_text_for_compare(value)
    type_code = canonical_type_name(field_type)

    if not type_code:
        return True, field_type

    if type_code == "text":
        return True, field_type

    if type_code == "email":
        compiled_email_regex = re.compile(email_regex)
        return compiled_email_regex.fullmatch(value_text) is not None, field_type

    if type_code == "numeric":
        return is_valid_decimal(value_text), field_type

    if type_code == "integer":
        return is_valid_integer(value_text), field_type

    if type_code == "decimal":
        return is_valid_decimal(value_text), field_type

    if type_code == "date":
        return is_valid_date(value_text), field_type

    if type_code == "boolean":
        return is_valid_boolean(value_text), field_type

    if type_code == "phone":
        return is_valid_phone(value_text), field_type

    if type_code == "alphanum":
        return is_valid_alphanum(value_text), field_type

    return True, field_type
