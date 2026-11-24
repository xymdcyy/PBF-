"""AO 登录守候：无登录框宽限与"最多自动填一次"（不碰 UIA）。

真实缺陷：已用缓存凭据登录时 AO 不弹框，旧逻辑 expect_login=True 会空等满
60s 再报"AO 自动登录失败"。另外凭据被拒时反复自动提交有账号锁定风险，
只允许自动尝试一次，之后必须由用户手动登录。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.ao_login as ao_login
from engine.ao_login import AoLogin


class _Clock:
    """替身时钟：sleep 推进时间，循环瞬时跑完。"""

    def __init__(self) -> None:
        self.t = 1000.0

    def time(self) -> float:
        return self.t

    def sleep(self, seconds: float) -> None:
        self.t += seconds


class EnsureTest(unittest.TestCase):
    def setUp(self):
        self.clock = _Clock()
        patcher = mock.patch.object(ao_login, "time", new=self.clock)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _login(self, boxes, fills=None):
        """boxes：每次 _find_login 依次返回的布尔值（用尽后重复最后一个）。"""
        login = AoLogin("u", "p")
        seq = list(boxes)
        calls = fills if fills is not None else []

        def find_login(timeout=0.5):
            return seq.pop(0) if len(seq) > 1 else seq[0]

        def fill():
            calls.append(1)
            return True

        login._find_login = find_login
        login._accept_insecure = lambda: False
        login._fill_and_submit = fill
        return login, calls

    def test_expect_login_no_box_passes_after_grace(self):
        login, fills = self._login([False])
        ok = login.ensure(timeout=60, expect_login=True, no_box_ok_after=30)
        self.assertTrue(ok)
        self.assertFalse(login.saw_login_box)
        self.assertEqual(fills, [])
        # 不能空等满 60s
        self.assertGreaterEqual(self.clock.t - 1000.0, 30)
        self.assertLess(self.clock.t - 1000.0, 60)

    def test_expect_login_without_grace_times_out(self):
        login, _ = self._login([False])
        self.assertFalse(login.ensure(timeout=10, expect_login=True))
        self.assertGreaterEqual(self.clock.t - 1000.0, 10)

    def test_no_expect_login_returns_quickly(self):
        login, _ = self._login([False])
        self.assertTrue(login.ensure(timeout=60, expect_login=False))
        self.assertLess(self.clock.t - 1000.0, 5)

    def test_box_then_gone_is_success(self):
        # 第一次检测到框 → 填写 → 再检测已消失
        login, fills = self._login([True, False])
        self.assertTrue(login.ensure(timeout=60, expect_login=True))
        self.assertEqual(len(fills), 1)
        self.assertTrue(login.saw_login_box)

    def test_box_never_clears_auto_fills_only_once(self):
        # 凭据被拒：绝不反复提交（账号锁定风险），交用户手动登录
        login, fills = self._login([True])
        ok = login.ensure(timeout=60, expect_login=True, max_attempts=1)
        self.assertFalse(ok)
        self.assertEqual(len(fills), 1)
        # 不会一路重试到超时：第二次见到框立即返回
        self.assertLess(self.clock.t - 1000.0, 20)


if __name__ == "__main__":
    unittest.main()
