from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_release_spec_is_valid_python() -> None:
    specification = PROJECT_ROOT / "excel-studio-release.spec"
    compile(specification.read_text(encoding="utf-8"), str(specification), "exec")


def test_release_spec_uses_current_product_entry_point() -> None:
    specification = (PROJECT_ROOT / "excel-studio-release.spec").read_text(encoding="utf-8")
    assert '"excel_studio" / "__main__.py"' in specification
    assert "_runtime_sources" not in specification
    assert "window_v6" not in specification
    assert "redesign" not in specification


def test_only_one_production_ui_tree_exists() -> None:
    ui_root = PROJECT_ROOT / "src" / "excel_studio" / "ui"
    expected = {
        "__init__.py",
        "base_page.py",
        "hierarchy_split_page.py",
        "hierarchy_value_selector.py",
        "main_window.py",
        "merge_page.py",
        "product_layout.py",
        "product_window.py",
        "split_page.py",
        "theme.py",
        "widgets.py",
    }
    actual = {path.name for path in ui_root.glob("*.py")}
    assert actual == expected
    assert not (ui_root / "pages").exists()
    assert not (ui_root / "redesign").exists()
