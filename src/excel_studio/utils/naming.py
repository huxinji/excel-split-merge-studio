from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


def sanitize_filename(value: str, fallback: str = "output", max_length: int = 180) -> str:
    cleaned = INVALID_WINDOWS_CHARS.sub("_", value).strip().rstrip(". ")
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        cleaned = fallback
    if cleaned.upper() in RESERVED_WINDOWS_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned[:max_length].rstrip(". ") or fallback


def sanitize_sheet_name(value: str, fallback: str = "Sheet") -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", value).strip("'")
    return (cleaned or fallback)[:31]


def render_naming_template(template: str, values: dict[str, Any]) -> str:
    now = datetime.now()
    defaults: dict[str, Any] = {
        "original_name": "workbook",
        "sheet_name": "Sheet",
        "split_field": "",
        "split_value": "",
        "part_no": 1,
        "total_parts": 1,
        "date": now.strftime("%Y%m%d"),
        "time": now.strftime("%H%M%S"),
        "datetime": now.strftime("%Y%m%d_%H%M%S"),
        "task_name": "task",
        "row_count": 0,
        "index": 1,
    }
    defaults.update(values)
    try:
        rendered = template.format_map(defaults)
    except (KeyError, ValueError) as error:
        raise ValueError(f"Invalid naming template: {error}") from error
    return sanitize_filename(rendered)


def unique_output_path(directory: Path, stem: str, suffix: str = ".xlsx") -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    candidate = directory / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    index = 2
    while True:
        candidate = directory / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1
