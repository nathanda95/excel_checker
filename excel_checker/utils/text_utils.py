import re
import unicodedata


def normalize_sheet_title(title: str) -> str:
    invalid = ['\\', '/', '*', '[', ']', ':', '?']
    for char in invalid:
        title = title.replace(char, "_")
    return title[:31]


def is_empty(value) -> bool:
    if value is None:
        return True

    if isinstance(value, str):
        cleaned = value.replace("\xa0", " ").strip()
        return cleaned == ""

    return False


def get_full_header_text(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def clean_header_name(value) -> str:
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ").strip()
    first_line = text.split("\n")[0].strip()
    return first_line


def make_human_header_name(value) -> str:
    text = clean_header_name(value)
    if not text:
        return ""
    return text.replace("_", " ").strip()


def normalize_text_for_compare(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def normalize_key(value: str) -> str:
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ").strip().lower()
    text = text.replace("*", " ")
    text = text.replace("_", " ")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_header_metadata(header_value):
    full_text = get_full_header_text(header_value)
    if not full_text:
        return {
            "field_name": "",
            "field_type": "",
            "max_length": None
        }

    lines = [line.strip() for line in full_text.splitlines() if line.strip()]
    field_name = lines[0] if lines else ""

    field_type = ""
    max_length = None

    for line in lines[1:]:
        lower_line = line.lower()

        if lower_line.startswith("type:"):
            field_type = line.split(":", 1)[1].strip()

        elif lower_line.startswith("longueur:"):
            raw = line.split(":", 1)[1].strip()
            match = re.search(r"\d+", raw)
            if match:
                try:
                    max_length = int(match.group())
                except ValueError:
                    max_length = None

    return {
        "field_name": field_name,
        "field_type": field_type,
        "max_length": max_length
    }


def header_contains_star(header_name: str) -> bool:
    return "*" in (header_name or "")


def parse_importance_label(value: str) -> str:
    text = normalize_key(value)

    if not text:
        return ""

    if "recommande" in text:
        return "recommended"

    if "obligatoire" in text:
        return "mandatory"

    return ""
