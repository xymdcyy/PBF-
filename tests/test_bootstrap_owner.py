"""Excel 归属记录、Quit 验证、登录守候参数（不碰 Excel/COM）。

背景：Quit 在存在外部 COM 引用时可能被忽略，旧代码只调 Quit 就记"已关闭"，
留下的无窗口僵尸实例会吞掉下一次导航表打开，表现为"AO 又没了"。
这里覆盖归属记录与"Quit 后必须确认进程退出"的纯逻辑。
"""
from __future__ import annotations

import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.bootstrap as bs


class OwnerRecordTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "logs" / "excel_owner.json"

    def tearDown(self):
        self._tmp.cleanup()

    def _write(self, pid=4321, created=1000, source=bs.OWNER_SOURCE_PBF):
        with mock.patch.object(bs, "_process_created", return_value=created):
            bs.write_excel_owner(pid, "导航.xlsb", source=source, path=self.path)

    def test_write_then_read(self):
        self._write()
        record = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(record["pid"], 4321)
        self.assertEqual(record["created"], 1000)
        self.assertEqual(record["source"], "pbf")
        with mock.patch.object(bs, "excel_pids", return_value=[4321]), \
             mock.patch.object(bs, "_process_created", return_value=1000):
            self.assertEqual(bs.read_excel_owner(self.path)["pid"], 4321)

    def test_pid_none_writes_nothing(self):
        bs.write_excel_owner(None, "导航.xlsb", path=self.path)
        self.assertFalse(self.path.exists())

    def test_process_gone_drops_record(self):
        self._write()
        with mock.patch.object(bs, "excel_pids", return_value=[]):
            self.assertIsNone(bs.read_excel_owner(self.path))
        self.assertFalse(self.path.exists())

    def test_pid_reused_drops_record(self):
        self._write(created=1000)
        with mock.patch.object(bs, "excel_pids", return_value=[4321]), \
             mock.patch.object(bs, "_process_created", return_value=2000):
            self.assertIsNone(bs.read_excel_owner(self.path))
        self.assertFalse(self.path.exists())

    def test_unknown_creation_time_still_trusted(self):
        # 取不到创建时间时只按 PID 匹配（Windows 权限受限等），不能因此误关
        self._write(created=None)
        with mock.patch.object(bs, "excel_pids", return_value=[4321]):
            self.assertIsNotNone(bs.read_excel_owner(self.path))

    def test_broken_json_is_ignored(self):
        self.path.parent.mkdir(parents=True)
        self.path.write_text("{not json", encoding="utf-8")
        self.assertIsNone(bs.read_excel_owner(self.path))

    def test_clear_only_matching_pid(self):
        self._write(pid=4321)
        bs.clear_excel_owner(9999, path=self.path)
        self.assertTrue(self.path.exists())
        bs.clear_excel_owner(4321, path=self.path)
        self.assertFalse(self.path.exists())


class RecordOwnerTest(unittest.TestCase):
    """启动/复用实例后登记归属：只有新 PID 才算本工具所有。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "excel_owner.json"
        patcher = mock.patch.object(bs, "owner_file", new=lambda: self.path)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self._tmp.cleanup)
        created = mock.patch.object(bs, "_process_created", return_value=1000)
        created.start()
        self.addCleanup(created.stop)
        pids = mock.patch.object(bs, "excel_pids", return_value=[100, 300])
        pids.start()
        self.addCleanup(pids.stop)

    def _engine(self, pid):
        return types.SimpleNamespace(pid=pid)

    def test_new_pid_is_pbf(self):
        source = bs._record_owner(self._engine(300), "nav", launched_pids={100})
        self.assertEqual(source, "pbf")
        self.assertEqual(bs.read_excel_owner()["source"], "pbf")

    def test_baseline_pid_is_reused(self):
        source = bs._record_owner(self._engine(100), "nav", launched_pids={100})
        self.assertEqual(source, "reused")
        self.assertEqual(bs.read_excel_owner()["source"], "reused")

    def test_existing_pbf_record_not_downgraded(self):
        # 上一轮本工具启动的 Excel 这一轮被复用：仍是我们的实例
        bs._record_owner(self._engine(300), "nav", launched_pids=set())
        source = bs._record_owner(self._engine(300), "nav", launched_pids={300})
        self.assertEqual(source, "pbf")
        self.assertEqual(bs.read_excel_owner()["source"], "pbf")


class QuitVerifiedTest(unittest.TestCase):
    def _engine(self, pid=100, quit_exc=None):
        def quit_():
            if quit_exc is not None:
                raise quit_exc
        return types.SimpleNamespace(pid=pid,
                                     _excel=types.SimpleNamespace(Quit=quit_))

    def _run(self, engine, wait_results, kill_calls):
        waits = iter(wait_results)

        def fake_wait(pid, timeout):
            return next(waits)

        with mock.patch.object(bs, "_wait_pid_exit", new=fake_wait), \
             mock.patch.object(bs, "reset_active_excel", new=lambda: None), \
             mock.patch.object(bs, "clear_excel_owner", new=lambda *a, **k: None), \
             mock.patch.object(bs, "kill_excel",
                               new=lambda pid=None: kill_calls.append(pid)):
            return bs.quit_excel_verified(engine, "测试")

    def test_quit_takes_effect(self):
        kills: list = []
        self.assertTrue(self._run(self._engine(), [True], kills))
        self.assertEqual(kills, [])

    def test_quit_ignored_then_killed(self):
        # 僵尸实例的典型表现：Quit 被忽略，必须验证并强制结束这个 PID
        kills: list = []
        self.assertTrue(self._run(self._engine(), [False, True], kills))
        self.assertEqual(kills, [100])

    def test_kill_failure_reported(self):
        kills: list = []
        self.assertFalse(self._run(self._engine(), [False, False], kills))
        self.assertEqual(kills, [100])

    def test_quit_exception_still_verifies(self):
        kills: list = []
        engine = self._engine(quit_exc=RuntimeError("COM 忙"))
        self.assertTrue(self._run(engine, [False, True], kills))
        self.assertEqual(kills, [100])

    def test_unknown_pid_never_kills_anything(self):
        kills: list = []
        self.assertFalse(self._run(self._engine(pid=None), [], kills))
        self.assertEqual(kills, [])


class LoginSecondsTest(unittest.TestCase):
    def test_default_when_missing(self):
        self.assertEqual(bs.launch_no_box_ok_sec({}), bs.LAUNCH_NO_BOX_OK_SEC)
        self.assertEqual(bs.launch_no_box_ok_sec({"login": {}}),
                         bs.LAUNCH_NO_BOX_OK_SEC)

    def test_valid_value(self):
        cfg = {"login": {"no_box_ok_sec": 45}}
        self.assertEqual(bs.launch_no_box_ok_sec(cfg), 45.0)

    def test_bool_falls_back(self):
        # YAML 误写 true：bool 是 int 子类，不能变成"只等 1 秒"
        cfg = {"login": {"no_box_ok_sec": True}}
        self.assertEqual(bs.launch_no_box_ok_sec(cfg), bs.LAUNCH_NO_BOX_OK_SEC)

    def test_invalid_and_nonpositive_fall_back(self):
        for raw in ("abc", 0, -5):
            cfg = {"login": {"no_box_ok_sec": raw}}
            self.assertEqual(bs.launch_no_box_ok_sec(cfg),
                             bs.LAUNCH_NO_BOX_OK_SEC, raw)


if __name__ == "__main__":
    unittest.main()
