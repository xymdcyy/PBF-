"""RunPackageRunner — executes BPC run packages via the PBF Pro run-log form.

From the recording, users open "PBF Pro运行包执行及日志查看清单.xlsb" from
navigator, click/execute the run package area (AA50 region or equivalent),
wait ~7-8 seconds, then click refresh to poll status.
"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass

from .excel import ExcelEngine, ExcelError

log = logging.getLogger(__name__)


@dataclass
class RunResult:
    status: str          # 已完成 / 失败 / 超时 / not_found
    start_time: str = ""
    end_time: str = ""
    log_id: str = ""
    message: str = ""


class RunPackageRunner:
    def __init__(self, engine: ExcelEngine, config: dict) -> None:
        self.engine = engine
        self.cfg = config

    def execute_and_wait(
        self,
        workbook_pattern: str = "PBF Pro运行包执行及日志查看清单*",
        wait_seconds: int = 15,
        poll_interval: float = 1.0,
    ) -> RunResult:
        wb = self.engine.get_workbook(workbook_pattern, wait_sec=30)
        sheet = self.engine.open_sheet(wb, self.cfg.get("sheet", "运行包执行及日志查看清单"))

        # Execute — the form exposes a macro named in config.execute_macro
        exec_macro = self.cfg.get("execute_macro")
        try:
            self.engine.run_macro(exec_macro, wb)
            log.info(f"Executed run package via {exec_macro}")
        except ExcelError as exc:
            # Fallback: maybe user clicks a button shape — try Trigger_PS generic
            log.warning(f"Custom macro failed ({exc}), trying generic Trigger_PS")
            try:
                self.engine.run_macro("Trigger_PS", wb)
            except ExcelError as exc2:
                return RunResult(status="execute_failed", message=f"{exc} | {exc2}")

        # Wait for completion, polling status cell
        status_cell = self.cfg.get("status_cell", "AA50")
        refresh_macro = self.cfg.get("refresh_macro", "Refresh_Sheet")
        start = time.time()
        last_val = None
        while time.time() - start < wait_seconds:
            time.sleep(poll_interval)
            try:
                self.engine.refresh(wb, macro_hint=refresh_macro)
                v = self.engine.read(sheet, status_cell)
                log.debug(f"poll {status_cell} -> {v}")
                if v != last_val:
                    last_val = v
                if "已完成" in str(v or "") or "成功" in str(v or ""):
                    return RunResult(status="已完成", message=str(v))
                if "失败" in str(v or "") or "ERROR" in str(v or "").upper():
                    return RunResult(status="失败", message=str(v))
            except ExcelError as exc:
                log.warning(f"poll err: {exc}")

        return RunResult(status="超时", message=f"last status={last_val} after {wait_seconds}s")
