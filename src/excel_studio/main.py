from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PySide6.QtCore import QTimer

from excel_studio.app import create_application
from excel_studio.feature_contract import build_feature_contract


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Excel Split & Merge Studio")
    parser.add_argument("--smoke-test", action="store_true", help="Open briefly and exit")
    parser.add_argument(
        "--self-test-report",
        type=Path,
        default=None,
        help="Write the packaged feature contract as JSON and exit",
    )
    parser.add_argument("--config-dir", type=Path, default=None, help="Override settings folder")
    return parser


def main(argv: list[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    application, window = create_application(argv=sys.argv, config_dir=arguments.config_dir)
    window.show()
    if arguments.self_test_report is not None:
        arguments.self_test_report.parent.mkdir(parents=True, exist_ok=True)
        arguments.self_test_report.write_text(
            json.dumps(build_feature_contract(window), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    if arguments.smoke_test or arguments.self_test_report is not None:
        QTimer.singleShot(250, application.quit)
    result = application.exec()
    if arguments.smoke_test:
        print("SMOKE_OK: commercial split/merge workspace opened")
    return result


if __name__ == "__main__":
    raise SystemExit(main())
