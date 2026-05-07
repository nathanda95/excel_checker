from collections import Counter

from excel_checker.utils.text_utils import clean_header_name, header_contains_star


def hex_to_rgb(hex_color: str):
    if not hex_color:
        return None

    hex_color = hex_color.strip().replace("#", "").upper()

    if len(hex_color) != 6:
        return None

    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    except ValueError:
        return None


def is_mandatory_marker_cell(cell, target_rgb=None, tolerance=40):
    fill = cell.fill
    if fill is None:
        return False

    pattern = fill.patternType or fill.fill_type
    if pattern != "solid":
        return False

    def match_color(r, g, b):
        if target_rgb is not None:
            tr, tg, tb = target_rgb
            return (
                abs(r - tr) <= tolerance and
                abs(g - tg) <= tolerance and
                abs(b - tb) <= tolerance
            )

        return r >= 200 and g >= 150 and b <= 170

    color_candidates = []

    if fill.fgColor is not None:
        color_candidates.append(fill.fgColor)
    if hasattr(fill, "start_color") and fill.start_color is not None:
        color_candidates.append(fill.start_color)
    if hasattr(fill, "bgColor") and fill.bgColor is not None:
        color_candidates.append(fill.bgColor)

    for color in color_candidates:
        if color.type == "rgb" and color.rgb:
            rgb = color.rgb.upper()
            if len(rgb) == 8:
                rgb = rgb[2:]

            try:
                r = int(rgb[0:2], 16)
                g = int(rgb[2:4], 16)
                b = int(rgb[4:6], 16)

                if match_color(r, g, b):
                    return True
            except Exception:
                pass

        elif color.type == "indexed":
            if target_rgb is None:
                return True

        elif color.type == "theme":
            if target_rgb is None:
                return True

    return False


def get_all_non_empty_columns(ws, header_row: int):
    columns = []
    seen = set()

    for col_idx in range(1, ws.max_column + 1):
        raw_header = ws.cell(row=header_row, column=col_idx).value
        header_name = clean_header_name(raw_header)
        if not header_name:
            continue

        key = (col_idx, header_name)
        if key not in seen:
            columns.append((col_idx, raw_header))
            seen.add(key)

    return columns


def find_required_columns_by_color(ws, header_row: int, mandatory_marker_row: int, target_rgb=None):
    candidates = []

    for col_idx in range(1, ws.max_column + 1):
        raw_header = ws.cell(row=header_row, column=col_idx).value
        header_name = clean_header_name(raw_header)
        if not header_name:
            continue

        marker_cell = ws.cell(row=mandatory_marker_row, column=col_idx)

        if is_mandatory_marker_cell(marker_cell, target_rgb=target_rgb):
            candidates.append((col_idx, raw_header, marker_cell.style_id))

    style_counter = Counter(style_id for _, _, style_id in candidates)

    mandatory_style_ids = {
        style_id for style_id, count in style_counter.items() if count >= 2
    }

    required = []
    seen = set()

    for col_idx in range(1, ws.max_column + 1):
        raw_header = ws.cell(row=header_row, column=col_idx).value
        header_name = clean_header_name(raw_header)
        if not header_name:
            continue

        marker_cell = ws.cell(row=mandatory_marker_row, column=col_idx)

        by_color = is_mandatory_marker_cell(marker_cell, target_rgb=target_rgb)
        by_style = marker_cell.style_id in mandatory_style_ids

        if by_color or by_style:
            key = (col_idx, header_name)
            if key not in seen:
                required.append((col_idx, raw_header))
                seen.add(key)

    return required


def find_required_columns_by_star(ws, header_row: int):
    required = []
    seen = set()

    for col_idx in range(1, ws.max_column + 1):
        raw_header = ws.cell(row=header_row, column=col_idx).value
        header_name = clean_header_name(raw_header)
        if not header_name:
            continue

        if header_contains_star(header_name):
            key = (col_idx, header_name)
            if key not in seen:
                required.append((col_idx, raw_header))
                seen.add(key)

    return required


def merge_column_lists(*lists_):
    merged = []
    seen = set()

    for lst in lists_:
        for col_idx, raw_header in lst:
            key = (col_idx, clean_header_name(raw_header))
            if key not in seen:
                merged.append((col_idx, raw_header))
                seen.add(key)

    return merged


def find_target_columns(ws, header_row: int, mandatory_marker_row: int, detect_by_color: bool, detect_by_star: bool, target_rgb=None):
    by_color = []
    by_star = []

    if detect_by_color:
        by_color = find_required_columns_by_color(ws, header_row, mandatory_marker_row, target_rgb=target_rgb)

    if detect_by_star:
        by_star = find_required_columns_by_star(ws, header_row)

    if detect_by_color or detect_by_star:
        return merge_column_lists(by_color, by_star)

    return get_all_non_empty_columns(ws, header_row)


def build_detection_mode_label(detect_by_color: bool, detect_by_star: bool, target_rgb=None):
    color_label = None
    if detect_by_color:
        color_label = "couleur personnalisée" if target_rgb is not None else "couleur auto"

    star_label = "*" if detect_by_star else None

    labels = [x for x in [color_label, star_label] if x]

    if labels:
        return " + ".join(labels)

    return "toutes les colonnes"
