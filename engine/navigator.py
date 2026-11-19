"""Navigator — version switching + form opening via PBF Pro navigation workbook.

From recording trace:
- Version switch: click AR45 (navigator sheet) → SAP AO "选择成员" popup
  → user picks target version → AR45 + filter-variable cell both update.
  Programmatic path (verified earlier): write member key to E4 + refresh.
- Form open: select a row in navigator → SelectionChange macro downloads .xlsb
  template into %LOCALAPPDATA%\\Temp\\sapaocache\\. We poll until workbook appears.
"""
from __future__ import annotations

import logging
import time
from typing import Optional

from .excel import ExcelEngine, ExcelError

log = logging.getLogger(__name__)

NAV_WB_PATTERN = "PBF Pro*全面预算系统流程导航*"


class Navigator:
    def __init__(self, engine: ExcelEngine, config: dict) -> None:
        self.engine = engine
        self.cfg = config

    def _nav_wb(self):
        return self.engine.get_workbook(NAV_WB_PATTERN, wait_sec=60)

    # ---------- version ----------

    def switch_version(self, member_key: str, display_name: Optional[str] = None) -> None:
        """Switch forecast version by writing the ZDCATEGRY cells + refresh.

        member_key example: "C03202611V1" (滚动预测2026年11月月初版)
        写完会校验 AR45 的显示文本；没生效就还原原值并报错（切版本必须人工确认）。
        """
        wb = self._nav_wb()
        sheet = self.engine.open_sheet(wb, self.cfg.get("nav_sheet", "PBF Pro - 全面预算系统流程导航"))
        wb.Activate()

        ctx_cell = self.cfg.get("version_context_cell")
        mem_cell = self.cfg.get("version_member_cell")
        disp_cell = self.cfg.get("version_display_cell", "AR45")
        targets = [c for c in (ctx_cell, mem_cell) if c]
        if not targets:
            raise ExcelError(
                "dimensions.yaml 未配置 navigator.version_context_cell，无法自动切版本；"
                "请在 AO 里手动选择版本")

        old = {c: self.engine.read(sheet, c) for c in targets}
        for c in targets:
            self.engine.write(sheet, c, member_key)

        try:
            self.engine.refresh(wb, macro_hint=self.cfg.get("refresh_macro"))
        except ExcelError:
            pass
        time.sleep(3.0)

        actual = str(self.engine.read(sheet, disp_cell) or "")
        if display_name and display_name not in actual:
            for c, v in old.items():
                self.engine.write(sheet, c, v)
            raise ExcelError(
                f"切版本未生效：{disp_cell} 仍为 '{actual}'（期望 {display_name}），已还原 {targets}。"
                f"请在 AO 中手动选择版本。")
        log.info(f"版本切换: {disp_cell} = {actual}")

    def current_version_key(self) -> str:
        wb = self._nav_wb()
        sheet = self.engine.open_sheet(wb, self.cfg.get("nav_sheet", "PBF Pro - 全面预算系统流程导航"))
        var_cell = self.cfg.get("version_context_cell") or "C20"
        v = self.engine.read(sheet, var_cell) or ""
        return str(v)

    # ---------- form opening ----------

    def open_form(self, workflow_row_col: str) -> str:
        """Select navigator cell that triggers AO to open a child form.

        workflow_row_col: cell like "AC63" — AC column holds 业务单元 names.
        Selection triggers VBA SelectionChange → downloads .xlsb template.
        Returns the name of the newly opened workbook (AO timestamped).
        """
        wb = self._nav_wb()
        sheet = self.engine.open_sheet(wb, self.cfg.get("nav_sheet", "PBF Pro - 全面预算系统流程导航"))

        wb.Activate()
        sheet.Activate()
        before = set(self.engine.workbook_names())

        # Step 1 — click the workflow row (business unit cell in AC column)
        self.engine.select(sheet, workflow_row_col)
        time.sleep(3.0)  # let step panel refresh

        # Step 2 — find the "打开" button in AT column of the step panel
        # Scan rows around the clicked row to find AT cell with "打开"
        try:
            row_num = int("".join(ch for ch in workflow_row_col if ch.isdigit()))
        except ValueError:
            row_num = None

        open_cell = None
        if row_num:
            for r in range(row_num - 2, row_num + 10):
                try:
                    v = self.engine.read(sheet, f"AT{r}")
                    if v and "打开" in str(v):
                        open_cell = f"AT{r}"
                        break
                except Exception:
                    pass

        if open_cell:
            log.info(f"Found '打开' at {open_cell}, clicking")
            self.engine.select(sheet, open_cell)
        else:
            # Fallback — try AT52 / AT53 generic action rows
            for try_cell in ("AT52", "AT53", "AT63"):
                try:
                    v = self.engine.read(sheet, try_cell)
                    if v and "打开" in str(v):
                        self.engine.select(sheet, try_cell)
                        open_cell = try_cell
                        break
                except Exception:
                    continue

        # Poll for new workbook
        deadline = time.time() + 30
        while time.time() < deadline:
            after = set(self.engine.workbook_names())
            new = after - before
            nav_names = {n for n in new if "导航" not in n}
            if nav_names:
                sorted_names = sorted(nav_names, key=lambda n: n, reverse=True)
                return sorted_names[0]
            time.sleep(0.5)
        raise ExcelError(f"No new workbook opened after selecting {workflow_row_col}")
