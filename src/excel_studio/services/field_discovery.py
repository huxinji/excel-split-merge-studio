from __future__ import annotations

from collections import OrderedDict
from typing import Any

from excel_studio.models.field_catalog import (
    FieldCatalog,
    FieldDiscoveryRequest,
    FieldSource,
)
from excel_studio.services.workbook_inspector import preview_workbook


def _valid_fields(columns: list[str]) -> list[str]:
    fields: list[str] = []
    seen: set[str] = set()
    for raw in columns:
        value = str(raw).strip()
        normalized = value.casefold()
        if not value or normalized.startswith("unnamed") or normalized in seen:
            continue
        seen.add(normalized)
        fields.append(value)
    return fields


def _sample_values(fields: list[str], rows: list[list[Any]]) -> dict[str, list[str]]:
    samples: dict[str, list[str]] = {}
    for column_index, field in enumerate(fields):
        values: OrderedDict[str, None] = OrderedDict()
        for row in rows:
            if column_index >= len(row) or row[column_index] is None:
                continue
            value = str(row[column_index]).strip()
            if value:
                values[value] = None
            if len(values) >= 3:
                break
        samples[field] = list(values)
    return samples


def discover_fields(
    requests: list[FieldDiscoveryRequest],
    header_mode: str,
    header_row: int,
    header_end_row: int,
    preview_rows: int = 20,
) -> FieldCatalog:
    catalog = FieldCatalog()
    for request in requests:
        try:
            preview = preview_workbook(
                request.path,
                request.sheet_name,
                header_mode,
                header_row,
                header_end_row,
                preview_rows,
            )
            fields = _valid_fields(preview.columns)
            catalog.sources.append(
                FieldSource(
                    path=request.path,
                    sheet_name=preview.sheet_name,
                    fields=fields,
                    samples=_sample_values(fields, preview.rows),
                )
            )
        except Exception as error:
            sheet = request.sheet_name or "<default>"
            catalog.warnings.append(f"{request.path.name}/{sheet}: {error}")
    return catalog
