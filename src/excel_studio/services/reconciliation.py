from __future__ import annotations

from excel_studio.models.result import ReconciliationResult


def reconcile_rows(
    input_rows: int, output_rows: int, excluded_rows: int = 0
) -> ReconciliationResult:
    result = ReconciliationResult(input_rows, output_rows, excluded_rows)
    if not result.is_balanced:
        result.warning = (
            f"Row reconciliation failed: input={input_rows}, output={output_rows}, "
            f"excluded={excluded_rows}"
        )
    return result
