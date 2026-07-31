from __future__ import annotations

from pathlib import Path

import pandas as pd

from excel_studio.utils.naming import sanitize_sheet_name, unique_output_path


def write_frame(
    frame: pd.DataFrame,
    directory: Path,
    stem: str,
    sheet_name: str = "Data",
    overwrite: bool = False,
) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    output = directory / f"{stem}.xlsx" if overwrite else unique_output_path(directory, stem)
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        frame.to_excel(writer, index=False, sheet_name=sanitize_sheet_name(sheet_name))
        worksheet = writer.sheets[sanitize_sheet_name(sheet_name)]
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))
    return output
