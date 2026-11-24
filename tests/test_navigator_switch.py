"""切版本：AO 刚登录时刷新未就绪，必须重写版本变量并轮询 AR45。

现场（2026-11-24 13:55）：新开导航表、登录刚成功，写入 C20/F20 后
RefreshAll 被吞掉，只读一次 AR45 得到空串就判失败。修复后：写入+刷新+
轮询直到生效，刷新异常不再静默，失败才还原。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.navigator as nav_mod
from engine.excel import ExcelError
from engine.navigator import Navigator

KEY = "C03202611V2"
DISPLAY = "滚动预测2026年11月财总版"
CFG = {
    "nav_sheet": "导航",
    "version_context_cell": "C20",
    "version_member_cell": "F20",
    "version_display_cell": "AR45",
    "version_switch_timeout_sec": 30,
}


class _Clock:
    def __init__(self) -> None:
        self.t = 1000.0

    def time(self) -> float:
        return self.t

    def sleep(self, seconds: float) -> None:
        self.t += seconds


class _FakeWb:
    def __init__(self) -> None:
        self.activated = False

    def Activate(self) -> None:      # noqa: N802
        self.activated = True


class _FakeEngine:
    def __init__(self, cells: dict, refresh_plan) -> None:
        self.cells = dict(cells)
        self.refresh_plan = refresh_plan
        self.refresh_calls = 0
        self.writes: list[tuple[str, object]] = []
        self.ao_calls: list[tuple] = []
        self.ao_fail: set[str] = set()
        self.version_var = ""
        self.reads: list[str] = []
        self.read_plan = None
        self.first_write_after_reads: int | None = None
        self.idle_waits: list[float] = []
        self.wb = _FakeWb()

    def wait_ui_idle(self, timeout_sec=30.0, poll_sec=1.0):
        self.idle_waits.append(timeout_sec)
        return True

    def get_workbook(self, pattern, wait_sec=60):
        return self.wb

    def open_sheet(self, wb, name):
        return object()

    def activate(self, wb):
        wb.Activate()

    def read(self, sheet, cell):
        self.reads.append(cell)
        if self.read_plan is not None:
            return self.read_plan(cell, len(self.reads))
        return self.cells.get(cell)

    def write(self, sheet, cell, value):
        if self.first_write_after_reads is None:
            self.first_write_after_reads = len(self.reads)
        self.writes.append((cell, value))
        self.cells[cell] = value

    def refresh(self, wb, macro_hint=None):
        self.refresh_calls += 1
        return self.refresh_plan(self.cells, self.refresh_calls)

    def ao_call(self, function, *args):
        self.ao_calls.append((function,) + args)
        if function in self.ao_fail:
            raise ExcelError(f"AO {function} 被拒")
        if function == "SAPSetVariable" and args[0] == "ZCCSO_ZDCATEGRY_010":
            self.version_var = args[1]
        elif function == "SAPGetVariable":
            # 真实 AO 的签名是（数据源, 变量, 属性）；参数写反时回读为空
            if len(args) == 3 and args[0] == "DS_PROCESSLIST":
                return self.version_var
            return -2146826273
        return 0


class SwitchVersionTest(unittest.TestCase):
    def setUp(self):
        self.clock = _Clock()
        patcher = mock.patch.object(nav_mod, "time", new=self.clock)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _nav(self, cells, plan):
        engine = _FakeEngine(cells, plan)
        return Navigator(engine, dict(CFG)), engine

    def test_applies_on_first_refresh(self):
        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.cells["AR45"], DISPLAY)
        self.assertEqual(engine.cells["C20"], KEY)
        self.assertEqual(engine.cells["F20"], KEY)
        self.assertEqual(engine.refresh_calls, 1)

    def test_waits_for_excel_idle_before_switching(self):
        # 登录后加载项自动刷新期间切版本会把 add-in 打崩：先等 UI 空闲
        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.idle_waits, [30.0])

    def test_rewrites_when_ao_init_overwrites_variables(self):
        # 第 1 次刷新模拟 AO 初始化把版本变量覆盖成空；第 3 次才真正生效
        def plan(cells, n):
            if n == 1:
                cells["C20"] = ""
                cells["F20"] = ""
            elif n >= 3 and cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.cells["AR45"], DISPLAY)
        self.assertEqual(engine.cells["C20"], KEY)
        self.assertEqual(engine.refresh_calls, 3)
        # 每次重试都重写 C20；成功后还会补写一次（AO 回显会覆盖 C20）
        self.assertGreaterEqual(engine.writes.count(("C20", KEY)), 3)

    def _set_variable_calls(self, engine):
        return [c for c in engine.ao_calls
                if c[0] == "SAPSetVariable" and c[2] == KEY]

    def test_waits_for_ao_ready_before_writing(self):
        # 登录后 AO 会先自动刷新；AR45 为空说明还没就绪，不能急着写/调 AO
        def read_plan(cell, n):
            if cell == "AR45" and n <= 2:
                return ""
            return engine.cells.get(cell)

        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        engine.read_plan = read_plan
        nav.switch_version(KEY, DISPLAY)
        self.assertGreaterEqual(engine.first_write_after_reads, 3)
        self.assertEqual(engine.cells["AR45"], DISPLAY)

    def test_sets_ds_processlist_version_variable(self):
        # 导航表「刷新状态」的等价操作：把版本写进 DS_PROCESSLIST 提示变量
        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertIn(
            ("SAPSetVariable", "ZCCSO_ZDCATEGRY_010", KEY,
             "INPUT_STRING", "DS_PROCESSLIST"),
            engine.ao_calls)
        self.assertIn(
            ("SAPExecuteCommand", "RefreshData", "DS_PROCESSLIST"),
            engine.ao_calls)
        # 设置变量要在 PauseVariableSubmit On/Off 之间
        pause_on = next(i for i, c in enumerate(engine.ao_calls)
                        if c == ("SAPExecuteCommand", "PauseVariableSubmit", "On"))
        pause_off = next(i for i, c in enumerate(engine.ao_calls)
                         if c == ("SAPExecuteCommand", "PauseVariableSubmit", "Off"))
        set_var = next(i for i, c in enumerate(engine.ao_calls)
                       if c[0] == "SAPSetVariable")
        self.assertLess(pause_on, set_var)
        self.assertLess(set_var, pause_off)
        self.assertEqual(engine.version_var, KEY)

    def test_reads_back_with_data_provider_first_after_resume(self):
        # 现场实证：SAPGetVariable 的签名是（数据源, 变量, 属性），写反会返回
        # 错误码；且 PauseVariableSubmit 期间回读恒为空，必须在 Off 之后读。
        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        get_var = next(i for i, c in enumerate(engine.ao_calls)
                       if c[0] == "SAPGetVariable")
        pause_off = next(i for i, c in enumerate(engine.ao_calls)
                         if c == ("SAPExecuteCommand", "PauseVariableSubmit", "Off"))
        self.assertLess(pause_off, get_var)
        self.assertEqual(
            engine.ao_calls[get_var],
            ("SAPGetVariable", "DS_PROCESSLIST",
             "ZCCSO_ZDCATEGRY_010", "VALUEASKEY"))

    def test_sets_time_variable_from_c19(self):
        # 与导航表 Refresh_Status 一致：期间变量也一起提交（C03 版本期间为空）
        def plan(cells, n):
            if cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({"C19": None}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertIn(
            ("SAPSetVariable", "ZCCSO_ZDTIME_048", "",
             "INPUT_STRING", "DS_PROCESSLIST"), engine.ao_calls)

    def test_version_variable_is_reapplied_each_attempt(self):
        def plan(cells, n):
            if n >= 3 and cells.get("C20") == KEY:
                cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(len(self._set_variable_calls(engine)), 3)
        self.assertEqual(engine.refresh_calls, 3)

    def test_ao_variable_errors_do_not_abort_switch(self):
        # 某些 AO 版本不允许程序设提示变量；AR45 仍是最终判据
        def plan(cells, n):
            cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        engine.ao_fail = {"SAPSetVariable"}
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.cells["AR45"], DISPLAY)

    def test_failure_restores_old_version_variable(self):
        def plan(cells, n):
            raise ExcelError("refresh failed: AO 未连接")

        nav, engine = self._nav(
            {"C20": "C03202611V1", "F20": "C03202611V1"}, plan)
        with self.assertRaises(ExcelError):
            nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.version_var, "C03202611V1")

    def test_rewrites_context_cell_after_ao_redisplay(self):
        # AO 的 AfterRedisplay 回调会把 C20 抄回旧值；成功后必须补写
        def plan(cells, n):
            cells["AR45"] = DISPLAY
            cells["C20"] = "滚动预测2026年11月月初版"

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.cells["C20"], KEY)
        self.assertEqual(engine.cells["F20"], KEY)

    def test_refresh_errors_are_retried(self):
        def plan(cells, n):
            if n <= 2:
                raise ExcelError("refresh failed: 被呼叫方拒绝接收呼叫")
            cells["AR45"] = DISPLAY

        nav, engine = self._nav({}, plan)
        nav.switch_version(KEY, DISPLAY)
        self.assertEqual(engine.cells["AR45"], DISPLAY)
        self.assertEqual(engine.refresh_calls, 3)

    def test_failure_restores_old_values_and_reports_refresh_errors(self):
        def plan(cells, n):
            raise ExcelError("refresh failed: AO 未连接")

        nav, engine = self._nav({"C20": "OLD", "F20": "OLD"}, plan)
        with self.assertRaises(ExcelError) as ctx:
            nav.switch_version(KEY, DISPLAY)
        msg = str(ctx.exception)
        self.assertIn("切版本未生效", msg)
        self.assertIn("重试", msg)
        self.assertIn("刷新失败", msg)
        self.assertEqual(engine.cells["C20"], "OLD")
        self.assertEqual(engine.cells["F20"], "OLD")
        self.assertGreaterEqual(engine.refresh_calls, 2)
        # 不是只试一次就放弃
        self.assertGreater(engine.writes.count(("C20", KEY)), 1)

    def test_no_display_name_skips_polling(self):
        nav, engine = self._nav({}, lambda cells, n: None)
        nav.switch_version(KEY, None)
        self.assertEqual(engine.cells["C20"], KEY)
        self.assertEqual(engine.refresh_calls, 1)

    def test_already_correct_version_still_sets_member_cell(self):
        def plan(cells, n):
            cells["AR45"] = DISPLAY

        nav, engine = self._nav({"AR45": DISPLAY}, plan)
        nav.switch_version(KEY, DISPLAY)
        # F20 是传给表单的成员，必须显式写入
        self.assertEqual(engine.cells["F20"], KEY)
        self.assertEqual(engine.refresh_calls, 1)


if __name__ == "__main__":
    unittest.main()
