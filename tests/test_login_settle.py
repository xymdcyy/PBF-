"""登录成功后要等 AO 自动刷新收尾，再调 SAP* 变量（不碰真实 Excel）。

现场（2026-11-24 15:23）：新开导航表 → 自动登录成功 → 1 秒后切版本，
加载项立刻弹「SAP BI 加载项已断开所有数据源连接」+「未将对象引用设置到
对象的实例」，之后所有数据源都不可用。原因是登录框关闭后加载项会自动
刷新所有数据源（实测占住 UI 线程 30s 级），刷新没完就调 SAPSetVariable
会把它打崩。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.bootstrap as bs


class _FakeIdleEngine:
    login_settled = False

    def __init__(self) -> None:
        self.calls: list[float] = []

    def wait_ui_idle(self, timeout_sec: float = 120.0, poll_sec: float = 1.0) -> bool:
        self.calls.append(timeout_sec)
        return True


class WaitAoLoginIdleTest(unittest.TestCase):
    def test_uses_config_times_and_marks_settled(self):
        engine = _FakeIdleEngine()
        cfg = {"login": {"settle_start_sec": 2, "idle_timeout_sec": 7}}
        with mock.patch.object(bs, "_sleep_cancellable") as sleep:
            bs._wait_ao_login_idle(engine, cfg)
        sleep.assert_called_once_with(2.0)
        self.assertEqual(engine.calls, [7.0])
        self.assertTrue(engine.login_settled)

    def test_defaults_when_config_missing(self):
        engine = _FakeIdleEngine()
        with mock.patch.object(bs, "_sleep_cancellable") as sleep:
            bs._wait_ao_login_idle(engine, {})
        sleep.assert_called_once_with(float(bs.LOGIN_SETTLE_START_SEC))
        self.assertEqual(engine.calls, [float(bs.LOGIN_IDLE_TIMEOUT_SEC)])

    def test_timeout_is_logged_not_raised(self):
        class NeverIdle(_FakeIdleEngine):
            def wait_ui_idle(self, timeout_sec=120.0, poll_sec=1.0):
                self.calls.append(timeout_sec)
                return False

        engine = NeverIdle()
        with mock.patch.object(bs, "_sleep_cancellable"), \
             self.assertLogs("engine.bootstrap", level="WARNING") as logs:
            bs._wait_ao_login_idle(engine, {"login": {"idle_timeout_sec": 3}})
        self.assertTrue(any("超时" in line for line in logs.output))


class SleepCancellableTest(unittest.TestCase):
    def test_cancel_interrupts_wait(self):
        from engine.cancel import CancelledError

        with mock.patch.object(bs, "raise_if_cancelled",
                               side_effect=CancelledError("取消")), \
             mock.patch.object(bs.time, "sleep") as sleep:
            with self.assertRaises(CancelledError):
                bs._sleep_cancellable(30)
        sleep.assert_not_called()

    def test_zero_or_negative_returns_immediately(self):
        with mock.patch.object(bs.time, "sleep") as sleep:
            bs._sleep_cancellable(0)
        sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
