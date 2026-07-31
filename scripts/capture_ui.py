from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PySide6.QtCore import QTimer

from excel_studio.app import create_application


def main() -> int:
    parser = argparse.ArgumentParser(description="Capture production UI reference images")
    parser.add_argument("--output-dir", type=Path, default=Path("docs/screenshots"))
    parser.add_argument("--source", type=Path, default=Path("examples/sales-example.xlsx"))
    arguments = parser.parse_args()
    output = arguments.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    application, window = create_application(
        argv=sys.argv,
        config_dir=output / ".capture-config",
    )
    window.resize(1366, 768)
    window.show()
    source = arguments.source.resolve()
    if source.is_file():
        window.split_page.load_path(source)

    def capture_split() -> None:
        window.grab().save(str(output / "split-field-workspace.png"))
        window.switch_workspace("merge")
        QTimer.singleShot(500, capture_merge)

    def capture_merge() -> None:
        window.grab().save(str(output / "merge-workspace.png"))
        application.quit()

    QTimer.singleShot(2500, capture_split)
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
