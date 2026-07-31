from __future__ import annotations

import argparse
import base64
import hashlib
import io
import zipfile
from pathlib import Path

VERSION = "3.0.0"


def build_payload(source_root: Path) -> tuple[str, str]:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted((source_root / "excel_studio").rglob("*.py")):
            if ".patch-backup" in path.name or "__pycache__" in path.parts:
                continue
            archive.write(path, path.relative_to(source_root).as_posix())
    payload = buffer.getvalue()
    return base64.b85encode(payload).decode("ascii"), hashlib.sha256(payload).hexdigest()


def launcher_source(payload: str, digest: str) -> str:
    return f'''# -*- coding: utf-8 -*-
"""Excel Split & Merge Studio {VERSION} portable source launcher.

Run with Python 3.11+ after installing the dependencies listed in requirements.txt.
The embedded application source is verified and extracted to a versioned local cache.
"""
from __future__ import annotations

import base64
import hashlib
import os
import shutil
import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

APP_VERSION = {VERSION!r}
PAYLOAD_SHA256 = {digest!r}
PAYLOAD = {payload!r}


def _cache_base() -> Path:
    if sys.platform.startswith("win"):
        root = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    elif sys.platform == "darwin":
        root = str(Path.home() / "Library" / "Caches")
    else:
        root = os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
    return Path(root) / "ExcelSplitMergeStudio" / "portable"


def _writable_cache_base() -> Path:
    preferred = _cache_base()
    try:
        preferred.mkdir(parents=True, exist_ok=True)
        return preferred
    except OSError:
        fallback = Path(tempfile.gettempdir()) / "ExcelSplitMergeStudio" / "portable"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def _prepare_source() -> Path:
    raw = base64.b85decode(PAYLOAD.encode("ascii"))
    actual = hashlib.sha256(raw).hexdigest()
    if actual != PAYLOAD_SHA256:
        raise RuntimeError("Embedded source integrity check failed / 内置源码完整性校验失败")
    target = _writable_cache_base() / f"{{APP_VERSION}}-{{PAYLOAD_SHA256[:12]}}"
    marker = target / ".payload.sha256"
    if marker.is_file() and marker.read_text(encoding="ascii").strip() == PAYLOAD_SHA256:
        return target
    staging = Path(tempfile.mkdtemp(prefix="excel-studio-", dir=target.parent))
    try:
        with zipfile.ZipFile(BytesIO(raw)) as archive:
            archive.extractall(staging)
        (staging / ".payload.sha256").write_text(PAYLOAD_SHA256, encoding="ascii")
        if target.exists():
            shutil.rmtree(target)
        staging.replace(target)
    finally:
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)
    return target


def main() -> int:
    source = _prepare_source()
    sys.path.insert(0, str(source))
    try:
        from excel_studio.main import main as application_main
    except ModuleNotFoundError as error:
        missing = error.name or "dependency"
        print(
            "Missing dependency: " + missing + "\\n"
            "Install project requirements first: python -m pip install -r requirements.txt\\n\\n"
            "缺少运行依赖：" + missing + "\\n"
            "请先安装项目依赖：python -m pip install -r requirements.txt",
            file=sys.stderr,
        )
        return 2
    return int(application_main())


if __name__ == "__main__":
    raise SystemExit(main())
'''


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the portable single-file Python launcher")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "Excel_Split_Merge_Studio.py",
    )
    arguments = parser.parse_args()
    project = Path(__file__).resolve().parents[1]
    payload, digest = build_payload(project / "src")
    arguments.output.write_text(launcher_source(payload, digest), encoding="utf-8", newline="\n")
    print(f"Created {arguments.output} ({arguments.output.stat().st_size:,} bytes)")
    print(f"Payload SHA-256: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
