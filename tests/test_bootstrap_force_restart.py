"""ensure_excel_ready 的 force_restart 分支：定位不到旧 PID 时不许复用死会话。

原缺陷：AO 会话失效恢复时若旧 Excel PID 定位不到（PID 记录丢失/进程
易主），只打告警就落进「复用已有 Excel」路径——Excel 是单实例，Dispatch
连回的正是那个死会话实例，重启等于没重启，后续宏照旧报 0x800A9C68。
修复后：导航表还开着（死会话实例还活着）→ 直接报错让用户手动关；
导航表没开（旧实例确实退了）→ 按新会话正常启动。
"""
from __future__ import annotations

import sys
import types
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.bootstrap as bs
from engine.bootstrap import ensure_excel_ready
from engine.excel import ExcelError

FLOW_CFG = {
    "navigator": {
        "workbook_pattern": "PBF Pro - 全面预算系统流程导航*",
        "source": "file",
        # 三个用例都不该走到「打开导航表」之后；真走到时让它报路径错误
        "launch_file": "这个文件不该存在.xlsx",
    },
}


class _FakeExcelEngine:
    """替身 ExcelEngine：只暴露 force_restart 探测/复用路径用到的方法。"""

    def __init__(self, findable: bool = True) -> None:
        self._findable = findable

    def find_workbook(self, pattern):
        if not self._findable:
            raise ExcelError(f"没找到匹配 {pattern} 的工作簿")
        return types.SimpleNamespace(Name="PBF Pro - 全面预算系统流程导航.xlsb")

    def workbook_count(self) -> int:
        return 1    # 复用路径探测用：非 0 就不会去 Quit


def _fake_engine_factory(findable: bool = True):
    return lambda *a, **k: _FakeExcelEngine(findable=findable)


class ForceRestartTests(unittest.TestCase):
    def test_located_pid_is_killed_then_restart_proceeds(self) -> None:
        """stale_pid 还在运行 → 只杀它，然后继续走重启路径。"""
        killed: list = []
        pids = iter([[111], [], [], [], []])   # kill 之后进程消失
        with mock.patch.object(bs, "excel_pids", new=lambda: next(pids, [])), \
             mock.patch.object(bs, "kill_excel", new=lambda pid=None: killed.append(pid)), \
             mock.patch.object(bs, "ExcelEngine", new=_fake_engine_factory(findable=False)):
            with self.assertRaises(ExcelError) as ctx:
                ensure_excel_ready(FLOW_CFG, kill_stale=False, stale_pid=111,
                                   force_restart=True)
        self.assertEqual(killed, [111])
        # 走到了「打开导航表」的路径报错（launch_file 无效），而不是复用旧实例
        self.assertIn("路径无效", str(ctx.exception))

    def test_unlocated_pid_with_nav_open_raises(self) -> None:
        """旧 PID 不在但导航表还开着 → 死会话实例还活着，必须报错让用户手关。"""
        with mock.patch.object(bs, "excel_pids", new=lambda: [222]), \
             mock.patch.object(bs, "ExcelEngine", new=_fake_engine_factory(findable=True)):
            with self.assertRaises(ExcelError) as ctx:
                ensure_excel_ready(FLOW_CFG, kill_stale=False, stale_pid=999,
                                   force_restart=True)
        self.assertIn("手动关闭", str(ctx.exception))
        self.assertIn("导航表仍开着", str(ctx.exception))

    def test_unlocated_pid_without_nav_starts_fresh(self) -> None:
        """旧 PID 不在且导航表没开（旧实例已退出）→ 按新会话正常走启动路径。"""
        pids = iter([[222], [222], [222], [222], [222]])
        with mock.patch.object(bs, "excel_pids", new=lambda: next(pids, [222])), \
             mock.patch.object(bs, "ExcelEngine", new=_fake_engine_factory(findable=False)):
            with self.assertRaises(ExcelError) as ctx:
                ensure_excel_ready(FLOW_CFG, kill_stale=False, stale_pid=999,
                                   force_restart=True)
        # 复用失败后继续走到「打开导航表」的路径报错；不能是「手动关闭」
        self.assertIn("路径无效", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
