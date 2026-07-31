# Packaging / 打包

## Windows

```powershell
python -m pip install -e ".[dev,legacy]"
python scripts/build_single_file.py
python -m PyInstaller --noconfirm --clean excel-studio-release.spec
python scripts/verify_packaged_app.py dist/ExcelSplitMergeStudio.exe
```

Compile `installer/windows/ExcelSplitMergeStudio.iss` with Inno Setup to create the installer.

## macOS

Run the build on macOS; PyInstaller does not cross-compile macOS applications from Windows.

```bash
python -m pip install -e ".[dev,legacy]"
python -m PyInstaller --noconfirm --clean excel-studio-release.spec
python scripts/verify_packaged_app.py \
  dist/ExcelSplitMergeStudio.app/Contents/MacOS/ExcelSplitMergeStudio
hdiutil create -volname "Excel Split & Merge Studio" \
  -srcfolder dist/ExcelSplitMergeStudio.app -ov -format UDZO \
  release/ExcelSplitMergeStudio-3.0.0-macOS.dmg
```

The release workflow performs both native builds, runs the packaged feature contract, creates
SHA-256 checksums, and publishes assets when a `v*` tag is pushed.

Windows 与 macOS 必须在对应系统原生构建。发布工作流会验证打包后的默认字段拆分入口、
全部拆分/合并模式、中英切换、固定底栏和响应式窗口，再上传 Release。
