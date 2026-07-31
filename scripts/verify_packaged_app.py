from __future__ import annotations

import argparse
import json
import subprocess
import tempfile
from pathlib import Path
from typing import Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify a packaged Excel Studio executable")
    parser.add_argument("executable", type=Path)
    parser.add_argument("--timeout", type=float, default=45.0)
    return parser.parse_args()


def _validate_contract(contract: dict[str, Any]) -> None:
    expected = {
        "shell": "commercial_step_workspace",
        "default_workspace": "split",
        "split_default_mode": "by_field",
        "split_field_selector_present": True,
        "split_field_panel_present": True,
        "hierarchy_multiselect": True,
        "hierarchy_max_levels": 3,
        "fixed_footer": True,
        "responsive_window": True,
    }
    mismatches = {
        key: (expected_value, contract.get(key))
        for key, expected_value in expected.items()
        if contract.get(key) != expected_value
    }
    if mismatches:
        raise RuntimeError(f"Packaged feature contract mismatch: {mismatches}")
    if set(contract.get("language_switch", [])) != {"zh_CN", "en_US"}:
        raise RuntimeError("Packaged language switch is incomplete")
    if set(contract.get("split_modes", [])) != {
        "by_field",
        "hierarchy",
        "fixed_rows",
        "by_parts",
        "by_sheet",
    }:
        raise RuntimeError("Packaged split modes are incomplete")
    if set(contract.get("merge_modes", [])) != set(contract.get("expected_merge_modes", [])):
        raise RuntimeError("Packaged merge modes are incomplete")


def main() -> int:
    arguments = parse_args()
    executable = arguments.executable.resolve()
    if not executable.is_file():
        raise RuntimeError(f"Packaged executable is missing: {executable}")
    with tempfile.TemporaryDirectory(prefix="excel-studio-verify-") as temporary:
        root = Path(temporary)
        report = root / "feature-contract.json"
        completed = subprocess.run(
            [
                str(executable),
                "--self-test-report",
                str(report),
                "--config-dir",
                str(root / "config"),
            ],
            check=False,
            timeout=arguments.timeout,
        )
        if completed.returncode != 0:
            raise RuntimeError(f"Packaged application exited with code {completed.returncode}")
        if not report.is_file():
            raise RuntimeError("Packaged application did not write a feature contract")
        contract: dict[str, Any] = json.loads(report.read_text(encoding="utf-8"))
        _validate_contract(contract)
    print(f"PACKAGED_APP_OK: {executable}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
