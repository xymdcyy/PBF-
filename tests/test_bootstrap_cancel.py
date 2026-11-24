"""启动阶段取消：不能在打开导航表/等 AO 登录时拖到 Web 端硬杀进程树。

现场（2026-11-24 15:07）：用户在新开导航表、AO 还在加载时点「取消」。
启动阶段没有取消检查点，60s 宽限后 Web 端 taskkill /T，连刚打开的导航表
Excel 一起被杀——用户看到的就是「AO 又没了」。修复：启动轮询和登录等待
都检查取消标志，CancelledError 不被 ExcelError 吞掉，优雅退出并保留 Excel。
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.bootstrap as bs
from engine import cancel
from engine.cancel import CancelledError


class EnsureReadyCancelTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.nav_file = Path(self._tmp.name) / "nav.xlsb"
        self.nav_file.write_bytes(b"")
        self.cancel_path = Path(self._tmp.name) / "run.cancel"
        cancel.set_cancel_file(self.cancel_path)
        self.addCleanup(cancel.set_cancel_file, None)

    def _cfg(self):
        return {
            "navigator": {
                "source": "file",
                "launch_file": str(self.nav_file),
                "workbook_pattern": "PBF Pro*全面预算系统流程导航*",
            },
        }

    def test_cancel_before_start_touches_nothing(self):
        self.cancel_path.write_text("", encoding="utf-8")
        with mock.patch.object(bs, "excel_pids") as pids, \
             mock.patch.object(bs.subprocess, "Popen") as popen:
            with self.assertRaises(CancelledError):
                bs.ensure_excel_ready(self._cfg(), wait_sec=5)
        pids.assert_not_called()        # 连 Excel 进程都没去碰
        popen.assert_not_called()

    def test_cancel_while_waiting_is_not_swallowed(self):
        # 第 1 次检查（入口）放行；第 2 次检查（打开后的轮询）抛取消。
        # 轮询循环只吞 ExcelError，取消必须能穿透，否则仍会被硬杀。
        calls = {"n": 0}

        def fake_check():
            calls["n"] += 1
            if calls["n"] >= 2:
                raise CancelledError("已按用户请求取消")

        with mock.patch.object(bs, "excel_pids", return_value=[]), \
             mock.patch.object(bs, "_start_ao_login", return_value=None), \
             mock.patch.object(bs, "raise_if_cancelled", side_effect=fake_check), \
             mock.patch.object(bs.ExcelEngine, "_real_excel_path",
                               return_value=r"C:\EXCEL.EXE"), \
             mock.patch.object(bs.ExcelEngine, "is_workbook_open",
                               return_value=False), \
             mock.patch.object(bs.time, "sleep", new=lambda s: None), \
             mock.patch.object(bs.subprocess, "Popen") as popen:
            with self.assertRaises(CancelledError):
                bs.ensure_excel_ready(self._cfg(), wait_sec=5)
        self.assertEqual(popen.call_count, 1)   # 已经打开了导航表，但保持不动


if __name__ == "__main__":
    unittest.main()
