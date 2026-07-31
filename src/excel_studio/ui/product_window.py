from __future__ import annotations

from PySide6.QtWidgets import QStackedWidget, QVBoxLayout, QWidget

from excel_studio.ui.hierarchy_split_page import HierarchySplitPage
from excel_studio.ui.main_window import MainWindow
from excel_studio.ui.merge_page import MergePage


class ProductMainWindow(MainWindow):
    """Main product shell using the hierarchy-capable split workspace."""

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)

        self.header = self._build_header()
        layout.addWidget(self.header)
        self.page_stack = QStackedWidget()
        self.split_page = HierarchySplitPage(self.i18n, self.settings)
        self.merge_page = MergePage(self.i18n, self.settings)
        self.page_stack.addWidget(self.split_page)
        self.page_stack.addWidget(self.merge_page)
        layout.addWidget(self.page_stack, 1)
        self.footer = self._build_footer()
        layout.addWidget(self.footer)
