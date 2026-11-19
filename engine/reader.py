"""Reader — read value snapshots from PBF forms.

Used to:
- preflight: capture existing values before writing
- verify: read values after refresh to confirm save landed
"""
from __future__ import annotations

from typing import Any

from .excel import ExcelEngine, ExcelError


class FormReader:
    def __init__(self, engine: ExcelEngine) -> None:
        self.engine = engine

    def read_cells(self, workbook_pattern: str, sheet_name: str, cells: list[str]) -> dict[str, Any]:
        wb = self.engine.get_workbook(workbook_pattern, wait_sec=15)
        sheet = self.engine.open_sheet(wb, sheet_name)
        out = {}
        for c in cells:
            try:
                out[c] = self.engine.read(sheet, c)
            except ExcelError as exc:
                out[c] = f"<err: {exc}>"
        return out
