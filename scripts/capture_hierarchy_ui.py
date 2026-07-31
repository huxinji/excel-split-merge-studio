from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook
from PySide6.QtCore import QTimer

from excel_studio.app import create_application
from excel_studio.models.operations import SplitMode


def _create_example(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "运营层级"
    sheet.append(["OPS", "大区", "片区", "网点", "负责人", "金额"])
    sheet.append(["中国运营", "华东", "共享片区", "浦东", "张敏", 12500])
    sheet.append(["中国运营", "华东", "上海", "浦西", "李杰", 9800])
    sheet.append(["中国运营", "华东", "浙江", "杭州", "王芳", 11300])
    sheet.append(["中国运营", "华南", "共享片区", "深圳", "刘洋", 14300])
    sheet.append(["中国运营", "华南", "广东", "广州", "周宁", 7600])
    workbook.save(path)
    workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Capture hierarchy split UI")
    parser.add_argument("--output-dir", type=Path, default=Path("docs/screenshots"))
    parser.add_argument(
        "--source",
        type=Path,
        default=Path("examples/hierarchy-example.xlsx"),
    )
    arguments = parser.parse_args()
    output = arguments.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    source = arguments.source.resolve()
    if not source.exists():
        _create_example(source)

    application, window = create_application(
        argv=sys.argv,
        config_dir=output / ".capture-hierarchy-config",
    )
    window.resize(1366, 768)
    window.show()
    page = window.split_page
    page.load_path(source)

    attempts = 0

    def configure() -> None:
        nonlocal attempts
        attempts += 1
        if page.current_preview is None or page.field_combo.findText("大区") < 0:
            if attempts < 60:
                QTimer.singleShot(200, configure)
            else:
                application.quit()
            return
        page.mode_selector.set_current(SplitMode.HIERARCHY.value)
        page.hierarchy_count_combo.setCurrentIndex(page.hierarchy_count_combo.findData(2))
        page.hierarchy_field_combos[0].setCurrentText("大区")
        page.hierarchy_field_combos[1].setCurrentText("片区")
        page.hierarchy_value_combos[0].set_selected_values(["华东", "华南"])
        QTimer.singleShot(1400, capture)

    def capture() -> None:
        window.grab().save(str(output / "hierarchy-split-workspace.png"))
        application.quit()

    QTimer.singleShot(200, configure)
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
