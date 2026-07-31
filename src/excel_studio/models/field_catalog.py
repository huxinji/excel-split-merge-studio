from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True, slots=True)
class FieldDiscoveryRequest:
    path: Path
    sheet_name: str | None


@dataclass(slots=True)
class FieldSource:
    path: Path
    sheet_name: str
    fields: list[str]
    samples: dict[str, list[str]] = field(default_factory=dict)


@dataclass(slots=True)
class FieldCatalog:
    sources: list[FieldSource] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
