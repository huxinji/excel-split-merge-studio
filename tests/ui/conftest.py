from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest
from PySide6.QtWidgets import QApplication

from excel_studio.app import create_application
from excel_studio.ui.main_window import MainWindow


@pytest.fixture
def product_window(qapp: QApplication, tmp_path: Path) -> Iterator[MainWindow]:
    del qapp
    _application, window = create_application(
        argv=[],
        config_dir=tmp_path / "settings",
    )
    window.show()
    yield window
    window.close()
