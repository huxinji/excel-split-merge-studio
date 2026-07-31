from __future__ import annotations

from pathlib import Path

APP_NAME = "Excel Split & Merge Studio"
APP_NAME_ZH = "Excel 拆分与合并工具"
ORGANIZATION_NAME = "Lucis WorkBuddy"
VERSION = "3.0.0"

SUPPORTED_EXCEL_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xls", ".xlsb"})
OPTIONAL_TEXT_EXTENSIONS = frozenset({".csv", ".tsv"})
SUPPORTED_EXTENSIONS = SUPPORTED_EXCEL_EXTENSIONS | OPTIONAL_TEXT_EXTENSIONS

DEFAULT_PREVIEW_ROWS = 20
DEFAULT_MAX_WORKERS = 2
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384

PROJECT_ROOT = Path(__file__).resolve().parents[3]
