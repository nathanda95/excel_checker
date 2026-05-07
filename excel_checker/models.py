from dataclasses import dataclass


@dataclass
class ColumnSelection:
    col_idx: int
    raw_header: str
    category: str = "mandatory"
