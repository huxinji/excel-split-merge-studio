from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from PySide6.QtCore import QStandardPaths

from excel_studio.config.constants import DEFAULT_MAX_WORKERS, DEFAULT_PREVIEW_ROWS


@dataclass(slots=True)
class AppSettings:
    language: str = "zh_CN"
    theme: str = "light"
    max_workers: int = DEFAULT_MAX_WORKERS
    preview_rows: int = DEFAULT_PREVIEW_ROWS
    recent_output_directories: list[str] | None = None

    def __post_init__(self) -> None:
        if self.recent_output_directories is None:
            self.recent_output_directories = []


class SettingsStore:
    """Small JSON settings store with an injectable base directory for tests."""

    def __init__(self, config_dir: Path | None = None) -> None:
        override = os.environ.get("EXCEL_STUDIO_CONFIG_DIR")
        if config_dir is not None:
            base_dir = config_dir
        elif override:
            base_dir = Path(override)
        else:
            qt_path = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
            base_dir = Path(qt_path) if qt_path else Path.home() / ".excel-studio"
        self._path = base_dir / "settings.json"

    @property
    def path(self) -> Path:
        return self._path

    def load(self) -> AppSettings:
        if not self._path.exists():
            return AppSettings()
        try:
            raw: dict[str, Any] = json.loads(self._path.read_text(encoding="utf-8"))
            allowed = {field.name for field in AppSettings.__dataclass_fields__.values()}
            return AppSettings(**{key: value for key, value in raw.items() if key in allowed})
        except (OSError, ValueError, TypeError):
            return AppSettings()

    def save(self, settings: AppSettings) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self._path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self._path)
