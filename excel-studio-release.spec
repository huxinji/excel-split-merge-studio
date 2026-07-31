# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path


project_root = Path(SPECPATH)
source_root = project_root / "src"

analysis = Analysis(
    [str(source_root / "excel_studio" / "__main__.py")],
    pathex=[str(source_root)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "pandas",
        "openpyxl",
        "xlsxwriter",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "IPython"],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="ExcelSplitMergeStudio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=sys.platform.startswith("win"),
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
)

if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="ExcelSplitMergeStudio.app",
        bundle_identifier="com.lucisworkbuddy.excelsplitmergestudio",
        info_plist={
            "CFBundleName": "Excel Split & Merge Studio",
            "CFBundleDisplayName": "Excel Split & Merge Studio",
            "CFBundleShortVersionString": "3.0.0",
            "CFBundleVersion": "3.0.0",
            "NSHighResolutionCapable": True,
        },
    )
