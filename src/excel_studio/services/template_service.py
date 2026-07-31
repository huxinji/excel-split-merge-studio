from __future__ import annotations

import json
from pathlib import Path
from typing import Any

SENSITIVE_KEYS = {"input_files", "input_paths", "output_directory", "directory", "paths"}


def _remove_sensitive_paths(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: _remove_sensitive_paths(item)
            for key, item in value.items()
            if key not in SENSITIVE_KEYS
        }
    if isinstance(value, list):
        return [_remove_sensitive_paths(item) for item in value]
    return value


class TemplateService:
    @staticmethod
    def save(path: Path, payload: dict[str, Any]) -> None:
        safe_payload = _remove_sensitive_paths(payload)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".tmp")
        temporary.write_text(
            json.dumps(safe_payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(path)

    @staticmethod
    def load(path: Path) -> dict[str, Any]:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("Template root must be an object")
        return payload
