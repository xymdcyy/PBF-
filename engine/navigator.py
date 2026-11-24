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

from .cancel import CancelledError, is_cancelled, raise_if_cancelled
from .excel import ExcelEngine, ExcelError

log = logging.getLogger(__name__)

NAV_WB_PATTERN = "PBF Pro*全面预算系统流程导航*"

# 切版本后等 AO 生效的总时长。AO 刚登录时服务端会话尚未就绪，写入的
# 版本变量会被随后的初始化覆盖、RefreshAll 也可能被拒——必须重写+重刷，
# 不能只等 3 秒读一次 AR45 就判失败。
VERSION_SWITCH_TIMEOUT_SEC = 60.0
VERSION_RETRY_INTERVAL_SEC = 2.0
VERSION_POLL_INTERVAL_SEC = 3.0
VERSION_POLL_STEP_SEC = 1.0

# 刚登录 AO 时会自动跑一遍 RefreshOnOpen；这期间 Excel 忙，AO 的
# Application.Run 调用会被长时间拒绝。先等 AR45 有值再开始切版本，
# 否则十几个 COM 调用全在重试，界面上看起来就是「卡死」。
VERSION_READY_TIMEOUT_SEC = 45.0
VERSION_READY_POLL_SEC = 3.0

# 导航表的版本回显链路（VBA Refresh_Status + sheet2 公式实证）：
#   Refresh_Status 把 C20 写进 DS_PROCESSLIST 的提示变量
#   ZCCSO_ZDCATEGRY_010（C19 → ZCCSO_ZDTIME_048），再 RefreshData
#   "DS_PROCESSLIST"；D10 = SAPGetVariable(...,"VALUEASKEY")、
#   AR45 = SAPGetMember("DS_PROCESSLIST","ZDCATEGRY="&...,"TEXT") 随之更新。
# 只写 C20/F20 + RefreshAll 不会碰这两个提示变量，AO 回显还会把 C20 抄回
# 旧值，所以 AR45 永远停在登录时的版本。
VERSION_DATA_PROVIDER = "DS_PROCESSLIST"
VERSION_PROMPT_KEY_VAR = "ZCCSO_ZDCATEGRY_010"
# Refresh_Status 还会把 C19（期间，C03 版本为空）写进 ZCCSO_ZDTIME_048；
# 只设版本变量不设期间变量，DS_PROCESSLIST 的提示状态和手动操作不一致。
VERSION_PROMPT_TIME_VAR = "ZCCSO_ZDTIME_048"
VERSION_TIME_CELL = "C19"


class Navigator:
    def __init__(self, engine: ExcelEngine, config: dict) -> None:
        self.engine = engine
        self.cfg = config

    def _nav_wb(self):
        return self.engine.get_workbook(NAV_WB_PATTERN, wait_sec=60)

    # ---------- version ----------

    def switch_version(self, member_key: str, display_name: Optional[str] = None) -> None:
        """Switch forecast version via the DS_PROCESSLIST prompt variable + refresh.

        member_key example: "C03202611V1" (滚动预测2026年11月月初版)
        写完会校验 AR45 的显示文本；没生效就还原原值并报错（切版本必须人工确认）。
        """
        # 登录框关闭后加载项会占住 Excel UI 线程自动刷新所有数据源（30s 级）。
        # 这期间调 SAPSetVariable 会把加载项打崩（「SAP BI 加载项已断开所有
        # 数据源连接」+ 空引用异常）。普通 Excel 调用当屏障：空闲时立即返回。
        idle_timeout = float(self.cfg.get("ao_idle_timeout_sec", 30.0))
        if not self.engine.wait_ui_idle(timeout_sec=idle_timeout):
            log.warning("切版本前等待 Excel 空闲超时（%.0fs），继续尝试", idle_timeout)

        wb = self._nav_wb()
        sheet = self.engine.open_sheet(wb, self.cfg.get("nav_sheet", "PBF Pro - 全面预算系统流程导航"))
        self.engine.activate(wb)

        ctx_cell = self.cfg.get("version_context_cell")
        mem_cell = self.cfg.get("version_member_cell")
        disp_cell = self.cfg.get("version_display_cell", "AR45")
        targets = [c for c in (ctx_cell, mem_cell) if c]
        if not targets:
            raise ExcelError(
                "dimensions.yaml 未配置 navigator.version_context_cell，无法自动切版本；"
                "请在 AO 里手动选择版本")

        log.info("切版本：目标 %s（%s）", display_name or member_key, member_key)
        self._wait_ao_ready(sheet, disp_cell)

        old = {c: self.engine.read(sheet, c) for c in targets}
        timeout = float(self.cfg.get("version_switch_timeout_sec",
                                     VERSION_SWITCH_TIMEOUT_SEC))
        deadline = time.time() + max(timeout, 0.0)
        ao_errors: list[str] = []
        refresh_errors: list[str] = []
        actual = ""
        attempts = 0

        # 写 C20/F20 → 设 DS_PROCESSLIST 版本提示变量 → 刷新 → 轮询 AR45；
        # 未生效就重来，直到超时。重写是必需的：AO 初始化/回显可能把刚写入
        # 的单元格覆盖掉。
        cancelled = False
        while True:
            if is_cancelled():
                # 取消时不能停在切换一半的版本上下文上，走下面的还原逻辑
                cancelled = True
                break
            attempts += 1
            log.info("切版本第 %d 次：写入 %s = %s，随后设置 AO 版本变量并刷新",
                     attempts, targets, member_key)
            for c in targets:
                self.engine.write(sheet, c, member_key)

            # 只写单元格不会驱动 AO：必须像「刷新状态」按钮那样把版本（和期间）
            # 写进 DS_PROCESSLIST 的提示变量。设失败只记录，AR45 仍是最终判据。
            ao_err = self._apply_version_variable(sheet, member_key, display_name)
            if ao_err:
                ao_errors.append(f"第{attempts}次: {ao_err}")
                log.warning("AO 设置版本变量失败（第 %d 次）：%s", attempts, ao_err)

            try:
                self.engine.refresh(wb, macro_hint=self.cfg.get("refresh_macro"))
            except ExcelError as exc:
                # 以前静默吞掉，日志里看不出"为什么 AR45 是空的"
                refresh_errors.append(str(exc))
                log.warning("切版本刷新失败（第 %d 次，AO 可能尚未就绪）：%s",
                            attempts, exc)

            actual = self._poll_display(sheet, disp_cell, display_name)
            if not display_name or display_name in actual:
                # AO 回显回调会把 C20 抄回 D10 的旧值；后续步骤按钮读的是
                # C20（Cells(20,3)），刷完再补写一次确保后续流程拿到目标版本。
                for c in targets:
                    self.engine.write(sheet, c, member_key)
                log.info(f"版本切换: {disp_cell} = {actual}")
                return
            if time.time() >= deadline:
                break
            log.info("版本尚未生效（%s='%s'，期望 %s），%.0fs 后重写并刷新（第 %d 次）",
                     disp_cell, actual, display_name,
                     VERSION_RETRY_INTERVAL_SEC, attempts + 1)
            time.sleep(VERSION_RETRY_INTERVAL_SEC)

        # 失败：还原单元格/版本提示变量并尽力刷新，回到切换前的上下文
        for c, v in old.items():
            try:
                self.engine.write(sheet, c, v)
            except ExcelError as exc:
                log.warning("还原 %s 失败（忽略）：%s", c, exc)
        self._restore_version_variable(old.get(ctx_cell) if ctx_cell else "")
        try:
            self.engine.refresh(wb, macro_hint=self.cfg.get("refresh_macro"))
        except ExcelError as exc:
            log.debug("还原版本变量后刷新失败（忽略）：%s", exc)

        if cancelled:
            raise CancelledError(
                "已按用户请求取消（切版本阶段，已还原原版本上下文）")

        detail = ""
        if refresh_errors:
            detail = (f"；刷新失败 {len(refresh_errors)} 次"
                      f"（最后一次：{refresh_errors[-1]}）")
        if ao_errors:
            detail += f"；AO 设置版本失败（最后一次：{ao_errors[-1]}）"
        raise ExcelError(
            f"切版本未生效：重试 {attempts} 次后 {disp_cell} 仍为 '{actual}'"
            f"（期望 {display_name}）{detail}，已还原 {targets}。"
            f"请确认 AO 已连接 BPC 后重试，或在 AO 中手动选择版本。")

    # ---------- version: AO prompt variable ----------

    def _version_provider(self) -> str:
        return str(self.cfg.get("version_data_provider", VERSION_DATA_PROVIDER))

    def _version_key_variable(self) -> str:
        return str(self.cfg.get("version_prompt_key_var", VERSION_PROMPT_KEY_VAR))

    def _wait_ao_ready(self, sheet, disp_cell: str) -> bool:
        """等 AO 登录后的自动刷新跑完（AR45 有值），避免后续 AO 调用全被拒。"""
        timeout = float(self.cfg.get("ao_ready_timeout_sec",
                                     VERSION_READY_TIMEOUT_SEC))
        deadline = time.time() + max(timeout, 0.0)
        log.info("切版本：等待 AO 就绪（读取 %s，最多 %.0fs）…", disp_cell, timeout)
        last_err = ""
        while True:
            raise_if_cancelled()
            try:
                value = self.engine.read(sheet, disp_cell)
            except ExcelError as exc:
                last_err = str(exc)
                value = None
            if value is not None and str(value).strip():
                log.info("AO 已就绪：%s = %r", disp_cell, str(value))
                return True
            if time.time() >= deadline:
                break
            time.sleep(VERSION_READY_POLL_SEC)
        log.warning("等待 AO 就绪超时（%s 为空或读取失败%s），仍继续尝试切版本",
                    disp_cell, f"：{last_err}" if last_err else "")
        return False

    def _apply_version_variable(self, sheet, member_key: str,
                                display_name: Optional[str] = None) -> Optional[str]:
        """把版本/期间写进 DS_PROCESSLIST 的提示变量并刷新；成功返回 None。

        等价于用户选好版本后点「刷新状态」（导航表 Refresh_Status）：
        设置变量前后暂停提交，避免每写一个变量就触发一次服务端请求。
        注意 AO 的 SAPGetVariable 参数顺序是（数据源, 变量, 属性），而且
        PauseVariableSubmit 期间回读恒为空——回读必须放在恢复提交之后。
        """
        dp = self._version_provider()
        var = self._version_key_variable()
        time_var = self._version_time_variable()
        time_cell = self._version_time_cell()
        errors: list[str] = []
        log.info("设置 AO 版本变量：%s.%s = %s（%s 取 %s）",
                 dp, var, member_key, time_var, time_cell)

        def call(*args):
            try:
                return self.engine.ao_call(*args)
            except ExcelError as exc:
                errors.append(f"{args[0]}: {exc}")
                log.debug(f"AO 调用失败（{args[0]}）：{exc}")
                return None

        try:
            time_value = self.engine.read(sheet, time_cell)
        except ExcelError as exc:
            time_value = ""
            log.debug(f"读取 {time_cell} 失败（按空处理）: {exc}")

        call("SAPSetRefreshBehaviour", "Off")
        call("SAPExecuteCommand", "PauseVariableSubmit", "On")
        set_ok = False
        try:
            if call("SAPSetVariable", var, member_key,
                    "INPUT_STRING", dp) is not None:
                set_ok = True
            # 与 Refresh_Status 一致：期间变量一起提交（C03 版本对应空期间）
            call("SAPSetVariable", time_var, str(time_value or ""),
                 "INPUT_STRING", dp)
        finally:
            call("SAPExecuteCommand", "PauseVariableSubmit", "Off")
            call("SAPSetRefreshBehaviour", "On")

        if not set_ok:
            errors.append(f"SAPSetVariable {var} 未生效")
        else:
            got = str(call("SAPGetVariable", dp, var, "VALUEASKEY") or "")
            log.info(f"AO 版本变量: {dp}.{var}={member_key} → 回读 {got!r}")
            if member_key not in got and not (display_name and display_name in got):
                errors.append(f"版本变量回读 {got!r}")

        try:
            self.engine.ao_call("SAPExecuteCommand", "RefreshData", dp)
        except ExcelError as exc:
            errors.append(f"RefreshData {dp}: {exc}")
        return "；".join(errors) if errors else None

    def _version_time_variable(self) -> str:
        return str(self.cfg.get("version_prompt_time_var",
                                VERSION_PROMPT_TIME_VAR))

    def _version_time_cell(self) -> str:
        return str(self.cfg.get("version_time_cell", VERSION_TIME_CELL))

    def _restore_version_variable(self, old_key: str) -> None:
        old_key = str(old_key or "").strip()
        if old_key[:3] not in ("C03", "C05"):
            return
        try:
            self.engine.ao_call(
                "SAPSetVariable", self._version_key_variable(), old_key,
                "INPUT_STRING", self._version_provider())
            self.engine.ao_call(
                "SAPExecuteCommand", "RefreshData", self._version_provider())
            log.info(f"已还原 AO 版本变量: {old_key}")
        except ExcelError as exc:
            log.warning("还原 AO 版本变量 %s 失败（忽略）：%s", old_key, exc)

    def _poll_display(self, sheet, disp_cell: str,
                      display_name: Optional[str]) -> str:
        """刷新后轮询显示单元格，AO 计算完成的时间不确定。"""
        actual = ""
        polls = max(1, int(VERSION_POLL_INTERVAL_SEC / VERSION_POLL_STEP_SEC))
        for _ in range(polls):
            time.sleep(VERSION_POLL_STEP_SEC)
            actual = str(self.engine.read(sheet, disp_cell) or "")
            if not display_name or display_name in actual:
                return actual
        return actual

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

        self.engine.activate(wb)
        self.engine.activate_sheet(sheet)
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
