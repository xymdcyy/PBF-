"""多 Excel 进程时按文件路径绑定实例（不碰真实 Excel/COM）。

现场（2026-11-24 14:49）：上一轮被杀掉的运行留下一个无窗口空实例，
它先注册进运行对象表；Dispatch 抓到的正是它。导航表其实开在另一个新
实例里，find_workbook 等 180s 也看不到导航表——表现为「等待导航表打开
超时」。修复：按文件路径从 ROT 精确绑定到开着导航表的实例。
"""
from __future__ import annotations

import sys
import types
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.excel as ex
from engine.excel import ExcelEngine, ExcelError

# 只用于路径比较，文件不必真实存在；按项目根拼，跨平台不写死盘符
NAV = str(Path(__file__).resolve().parent.parent / "PBF Pro - 全面预算系统流程导航.xlsb")


class RotPathTest(unittest.TestCase):
    def test_matches_same_file_regardless_of_case(self):
        with mock.patch.object(ExcelEngine, "_rot_display_names",
                               return_value=[NAV.upper()]):
            self.assertIsNotNone(ExcelEngine._rot_workbook_path(NAV))
            self.assertTrue(ExcelEngine.is_workbook_open(NAV))

    def test_other_workbooks_do_not_match(self):
        other = str(Path(NAV).with_name("其他表.xlsb"))
        with mock.patch.object(ExcelEngine, "_rot_display_names", return_value=[other]):
            self.assertFalse(ExcelEngine.is_workbook_open(NAV))

    def test_non_file_monikers_are_ignored(self):
        # ROT 里还有 "!{...}" 这类非文件条目，不能当成工作簿路径
        with mock.patch.object(ExcelEngine, "_rot_display_names",
                               return_value=["!{00024500-0000-0000-C000-000000000046}", NAV]):
            self.assertEqual(ExcelEngine._rot_workbook_path(NAV), NAV)


class FromRotTest(unittest.TestCase):
    def test_never_binds_when_not_open(self):
        # GetObject(未打开的文件) 会直接打开它——这里必须先查 ROT
        with mock.patch.object(ExcelEngine, "_rot_workbook_path", return_value=None), \
             mock.patch.object(ex.win32com.client, "GetObject") as get_object:
            self.assertIsNone(ExcelEngine._excel_from_rot(NAV))
        get_object.assert_not_called()

    def test_returns_application_of_open_workbook(self):
        app = types.SimpleNamespace(Version="16.0")
        wb = types.SimpleNamespace(Application=app)
        with mock.patch.object(ExcelEngine, "_rot_workbook_path", return_value=NAV), \
             mock.patch.object(ex.win32com.client, "GetObject", return_value=wb):
            self.assertIs(ExcelEngine._excel_from_rot(NAV), app)

    def test_rejects_wps(self):
        app = types.SimpleNamespace(Version="12.0")
        wb = types.SimpleNamespace(Application=app)
        with mock.patch.object(ExcelEngine, "_rot_workbook_path", return_value=NAV), \
             mock.patch.object(ex.win32com.client, "GetObject", return_value=wb):
            self.assertIsNone(ExcelEngine._excel_from_rot(NAV))


class ConnectWithWorkbookPathTest(unittest.TestCase):
    def setUp(self):
        ex.reset_active_excel()
        self.addCleanup(ex.reset_active_excel)

    def test_binds_via_rot_and_never_dispatches(self):
        app = types.SimpleNamespace(Version="16.0", Visible=False)
        with mock.patch.object(ExcelEngine, "_excel_from_rot", return_value=app), \
             mock.patch.object(ExcelEngine, "_dispatch_real_excel") as dispatch:
            engine = ExcelEngine(timeout_sec=1, workbook_path=NAV)
        self.assertIs(engine._excel, app)
        self.assertTrue(app.Visible)
        dispatch.assert_not_called()

    def test_not_open_raises_without_dispatching(self):
        # 未打开时必须快速失败让调用方轮询，不能回退到 Dispatch（可能连到空僵尸，
        # 更不能用 GetObject 把文件打开）
        with mock.patch.object(ExcelEngine, "_excel_from_rot", return_value=None), \
             mock.patch.object(ExcelEngine, "_dispatch_real_excel") as dispatch:
            with self.assertRaises(ExcelError):
                ExcelEngine(timeout_sec=1, workbook_path=NAV)
        dispatch.assert_not_called()

    def test_cached_instance_is_reused_after_binding(self):
        app = types.SimpleNamespace(Version="16.0", Visible=True)
        with mock.patch.object(ExcelEngine, "_excel_from_rot", return_value=app):
            first = ExcelEngine(timeout_sec=1, workbook_path=NAV)
        with mock.patch.object(ExcelEngine, "_dispatch_real_excel") as dispatch:
            second = ExcelEngine(timeout_sec=1)
        self.assertIs(second._excel, first._excel)
        dispatch.assert_not_called()


class ProbeVisibilityTest(unittest.TestCase):
    """探测连接不能把隐藏实例变可见（现场出现过空白 Excel 与导航表并排）。"""

    def setUp(self):
        ex.reset_active_excel()
        self.addCleanup(ex.reset_active_excel)

    def test_probe_leaves_hidden_instance_hidden(self):
        app = types.SimpleNamespace(Version="16.0", Visible=False)
        with mock.patch.object(ExcelEngine, "_dispatch_real_excel", return_value=app):
            engine = ExcelEngine(timeout_sec=1, make_visible=False)
        self.assertIs(engine._excel, app)
        self.assertFalse(app.Visible)

    def test_default_still_makes_instance_visible(self):
        app = types.SimpleNamespace(Version="16.0", Visible=False)
        with mock.patch.object(ExcelEngine, "_dispatch_real_excel", return_value=app):
            ExcelEngine(timeout_sec=1)
        self.assertTrue(app.Visible)


class _FakeLoginHandle:
    def __init__(self, outcome=True):
        self.outcome = outcome
        self.saw_box = False
        self.finished = True
        self.waited = False

    def wait(self, timeout=None):
        self.waited = True
        return self.outcome


class _FakeNavEngine:
    pid = 3752

    def __init__(self) -> None:
        self.find_calls = 0
        self._excel = types.SimpleNamespace(DisplayAlerts=True)

    def find_workbook(self, pattern):
        self.find_calls += 1
        return object()

    def ensure_visible(self):
        pass


class EnsureExcelReadyBindingTest(unittest.TestCase):
    """ensure_excel_ready：导航表已在另一个实例里开着时，不再重复启动。"""

    def setUp(self):
        import tempfile
        import engine.bootstrap as bs

        self.bs = bs
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.nav_file = Path(self._tmp.name) / "PBF Pro - 全面预算系统流程导航.xlsb"
        self.nav_file.write_bytes(b"")   # 只要求路径存在，不会真打开
        self.flow_cfg = {
            "navigator": {
                "source": "file",
                "launch_file": str(self.nav_file),
                "workbook_pattern": "PBF Pro*全面预算系统流程导航*",
            },
        }

    def _install_fakes(self, already_open: bool):
        bs = self.bs
        engine = _FakeNavEngine()
        calls: list[dict] = []

        def make_engine(timeout_sec=30.0, workbook_path=None, make_visible=True):
            calls.append({"timeout": timeout_sec, "workbook_path": workbook_path,
                          "make_visible": make_visible})
            # 第 1/2 步（3s 超时）都失败；第 4 步按路径绑定（5s）成功
            if workbook_path is not None and timeout_sec == 5:
                return engine
            raise ExcelError("未绑定")

        class FakeExcelEngine:
            _real_excel_path = staticmethod(lambda: r"C:\EXCEL.EXE")
            is_workbook_open = staticmethod(lambda p: already_open)

            def __new__(cls, *a, **k):
                return make_engine(*a, **k)

        handle = _FakeLoginHandle()
        patchers = [
            mock.patch.object(bs, "ExcelEngine", FakeExcelEngine),
            mock.patch.object(bs, "excel_pids", return_value=[10352]),
            mock.patch.object(bs, "read_excel_owner", return_value=None),
            mock.patch.object(bs, "clear_excel_owner", new=lambda *a, **k: None),
            mock.patch.object(bs, "_start_ao_login", return_value=handle),
            mock.patch.object(bs, "_ensure_visible", new=lambda e: None),
            mock.patch.object(bs, "_record_owner", return_value="reused"),
            mock.patch.object(bs.time, "sleep", new=lambda s: None),
        ]
        for p in patchers:
            p.start()
            self.addCleanup(p.stop)
        popen_patch = mock.patch.object(bs.subprocess, "Popen")
        popen = popen_patch.start()
        self.addCleanup(popen_patch.stop)
        return engine, calls, popen, handle

    def test_open_navigator_is_bound_not_relaunched(self):
        engine, calls, popen, handle = self._install_fakes(already_open=True)
        got = self.bs.ensure_excel_ready(self.flow_cfg, wait_sec=20)
        self.assertIs(got, engine)
        popen.assert_not_called()        # 文件已经开着，绝不能再启动一份
        self.assertTrue(handle.waited)
        # 最后一次构造必须是按路径绑定
        self.assertEqual(calls[-1]["workbook_path"], str(self.nav_file))
        # 探测性质的无路径 Dispatch 一律不能改可见性（防止多弹空白 Excel）
        self.assertTrue(all(c["make_visible"] is False for c in calls
                            if c["workbook_path"] is None))

    def test_not_open_launches_once(self):
        engine, calls, popen, handle = self._install_fakes(already_open=False)
        got = self.bs.ensure_excel_ready(self.flow_cfg, wait_sec=20)
        self.assertIs(got, engine)
        self.assertEqual(popen.call_count, 1)
        self.assertEqual(calls[-1]["workbook_path"], str(self.nav_file))
        self.assertTrue(all(c["make_visible"] is False for c in calls
                            if c["workbook_path"] is None))

    def test_fresh_login_waits_for_ao_refresh(self):
        # 登录框关闭后加载项还在自动刷新，直接切版本会把它打崩——必须等收尾
        engine, calls, popen, handle = self._install_fakes(already_open=False)
        handle.saw_box = True

        def settle_then_flag(e, cfg):
            e.login_settled = True       # 模拟真实屏障完成时打的标志

        with mock.patch.object(self.bs, "wait_ao_login_idle",
                               side_effect=settle_then_flag) as settle:
            self.bs.ensure_excel_ready(self.flow_cfg, wait_sec=20)
        settle.assert_called_once_with(engine, self.flow_cfg)
        # 顺序断言：屏障标志必须在 _mark_login 之后落盘——先 mark 后 settle，
        # 若反过来 mark 会把标志清回 False，run_flow 就会重复等一遍屏障
        self.assertTrue(getattr(engine, "login_settled", False))

    def test_reused_instance_is_not_recorded_as_owned(self):
        # 第 1 步直接复用了已开着的导航表：这个实例不是本次启动的，
        # 归属必须是 reused（或沿用已有 pbf 记录），不能新记成 pbf。
        bs = self.bs
        engine = _FakeNavEngine()

        def make_engine(timeout_sec=30.0, workbook_path=None, make_visible=True):
            return engine

        class FakeExcelEngine:
            _real_excel_path = staticmethod(lambda: r"C:\EXCEL.EXE")

            def __new__(cls, *a, **k):
                return make_engine(*a, **k)

        handle = _FakeLoginHandle()
        handle.saw_box = True       # 即便后台守候线程填过登录框，也不在复用路径等刷新
        settle_patch = mock.patch.object(bs, "wait_ao_login_idle")
        patches = [
            mock.patch.object(bs, "ExcelEngine", FakeExcelEngine),
            mock.patch.object(bs, "excel_pids", return_value=[engine.pid]),
            mock.patch.object(bs, "_start_ao_login", return_value=handle),
            mock.patch.object(bs, "_ensure_visible", new=lambda e: None),
            mock.patch.object(bs, "_record_owner", return_value="reused"),
        ]
        started = [p.start() for p in patches]
        settle = settle_patch.start()
        self.addCleanup(settle_patch.stop)
        for p in patches:
            self.addCleanup(p.stop)
        rec = started[-1]
        bs.ensure_excel_ready(self.flow_cfg, wait_sec=5)
        rec.assert_called_once()
        self.assertEqual(rec.call_args.kwargs.get("launched_pids"),
                         {engine.pid})
        settle.assert_not_called()   # 复用路径立即返回，不额外拖延


if __name__ == "__main__":
    unittest.main()
