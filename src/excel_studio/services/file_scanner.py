from __future__ import annotations

from collections.abc import Callable, Iterable
from pathlib import Path

from excel_studio.config.constants import OPTIONAL_TEXT_EXTENSIONS, SUPPORTED_EXCEL_EXTENSIONS
from excel_studio.models.task import ScanOptions
from excel_studio.models.workbook import FileRecord
from excel_studio.services.workbook_inspector import build_file_record


def collect_input_files(paths: Iterable[Path], options: ScanOptions) -> list[Path]:
    extensions = set(SUPPORTED_EXCEL_EXTENSIONS)
    if options.include_text_files:
        extensions.update(OPTIONAL_TEXT_EXTENSIONS)
    collected: dict[str, Path] = {}
    for input_path in paths:
        path = input_path.expanduser().resolve()
        candidates = path.rglob("*") if path.is_dir() and options.scan_subfolders else None
        if path.is_dir() and candidates is None:
            candidates = path.glob("*")
        if path.is_file():
            candidates = [path]
        if candidates is None:
            continue
        for candidate in candidates:
            if not candidate.is_file() or candidate.suffix.casefold() not in extensions:
                continue
            if options.ignore_temp_files and candidate.name.startswith("~$"):
                continue
            if options.ignore_hidden and candidate.name.startswith("."):
                continue
            folded_name = candidate.name.casefold()
            if options.name_contains and options.name_contains.casefold() not in folded_name:
                continue
            if options.name_excludes and options.name_excludes.casefold() in folded_name:
                continue
            collected[str(candidate).casefold()] = candidate
    return sorted(collected.values(), key=lambda item: str(item).casefold())


def scan_files(
    paths: Iterable[Path],
    options: ScanOptions,
    progress: Callable[[int, int, str], None] | None = None,
) -> list[FileRecord]:
    candidates = collect_input_files(paths, options)
    records: list[FileRecord] = []
    total = len(candidates)
    for index, path in enumerate(candidates, start=1):
        records.append(build_file_record(path))
        if progress is not None:
            progress(index, total, path.name)
    return records
