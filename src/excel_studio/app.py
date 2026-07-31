from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import QApplication

from excel_studio.config.constants import APP_NAME, ORGANIZATION_NAME, VERSION
from excel_studio.config.i18n import TranslationManager
from excel_studio.config.settings import SettingsStore
from excel_studio.ui.product_layout import (
    apply_accessible_contrast,
    prioritize_field_split_workflow,
    protect_header_legibility,
)
from excel_studio.ui.product_window import ProductMainWindow


def _product_font() -> QFont:
    available = set(QFontDatabase.families())
    candidates = (
        "Microsoft YaHei UI",
        "PingFang SC",
        "Noto Sans CJK SC",
        "Noto Sans SC",
        "Arial",
    )
    family = next((candidate for candidate in candidates if candidate in available), "Arial")
    return QFont(family, 10)


def create_application(
    argv: list[str] | None = None, config_dir: Path | None = None
) -> tuple[QApplication, ProductMainWindow]:
    existing = QApplication.instance()
    application = (
        existing
        if isinstance(existing, QApplication)
        else QApplication(argv if argv is not None else sys.argv)
    )
    application.setApplicationName(APP_NAME)
    application.setApplicationVersion(VERSION)
    application.setOrganizationName(ORGANIZATION_NAME)
    application.setStyle("Fusion")
    application.setFont(_product_font())

    store = SettingsStore(config_dir)
    settings = store.load()
    i18n = TranslationManager(settings.language)
    window = ProductMainWindow(settings, store, i18n)
    prioritize_field_split_workflow(window.split_page)
    protect_header_legibility(window)
    apply_accessible_contrast(window)
    return application, window
