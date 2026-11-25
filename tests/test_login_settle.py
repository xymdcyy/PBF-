"""登录成功后要等 AO 自动刷新收尾，再调 SAP* 变量（不碰真实 Excel）。

现场（2026-11-24 15:23）：新开导航表 → 自动登录成功 → 1 秒后切版本，
加载项立刻弹「SAP BI 加载项已断开所有数据源连接」+「未将对象引用设置到
对象的实例」，之后所有数据源都不可用。原因是登录框关闭后加载项会自动
刷新所有数据源（实测占住 UI 线程 30s 级），刷新没完就调 SAPSetVariable
会把它打崩。
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
import engine.orchestrator as orch


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
            bs.wait_ao_login_idle(engine, cfg)
        sleep.assert_called_once_with(2.0)
        self.assertEqual(engine.calls, [7.0])
        self.assertTrue(engine.login_settled)

    def test_defaults_when_config_missing(self):
        engine = _FakeIdleEngine()
        with mock.patch.object(bs, "_sleep_cancellable") as sleep:
            bs.wait_ao_login_idle(engine, {})
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
            bs.wait_ao_login_idle(engine, {"login": {"idle_timeout_sec": 3}})
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


class AwaitBackgroundLoginTest(unittest.TestCase):
    """orchestrator.await_background_login：分片可取消等待 + 成功后补收尾屏障。"""

    @staticmethod
    def _handle(finished=True, outcome=None, saw_box=False, wait=None):
        class _Handle:
            pass

        h = _Handle()
        h.finished = finished
        h.outcome = outcome
        h.saw_box = saw_box
        h.error = ""
        if wait is not None:
            h.wait = wait
        return h

    @staticmethod
    def _engine(handle=None, settled=False):
        class _Engine:
            pass

        e = _Engine()
        e.login_handle = handle
        e.login_settled = settled
        return e

    def test_confirmed_login_with_box_settles(self):
        # 后台线程填过登录框且成功：等 AO 自动刷新收尾（复用路径的补漏）
        handle = self._handle(finished=True, outcome=True, saw_box=True)
        engine = self._engine(handle)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle:
            self.assertTrue(orch.await_background_login(engine, {}, 5))
        settle.assert_called_once_with(engine, {})

    def test_already_settled_skips_barrier(self):
        # bootstrap 新启动路径已做过屏障：不重复等
        handle = self._handle(finished=True, outcome=True, saw_box=True)
        engine = self._engine(handle, settled=True)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle:
            self.assertTrue(orch.await_background_login(engine, {}, 5))
        settle.assert_not_called()

    def test_no_box_no_barrier(self):
        # 登录框从未出现（缓存凭据已登录）：没有自动刷新，不需要屏障
        handle = self._handle(finished=True, outcome=True, saw_box=False)
        engine = self._engine(handle)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle:
            self.assertTrue(orch.await_background_login(engine, {}, 5))
        settle.assert_not_called()

    def test_failed_login_returns_false(self):
        handle = self._handle(finished=True, outcome=False, saw_box=True)
        engine = self._engine(handle)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle:
            self.assertFalse(orch.await_background_login(engine, {}, 5))
        settle.assert_not_called()

    def test_unfinished_handle_waits_for_outcome(self):
        # 线程还没出结果：等它结束再判定（不能拿"已启动"当登录成功）
        def wait(timeout=None):
            handle.finished = True
            handle.outcome = True
            return handle.outcome

        handle = self._handle(finished=False, saw_box=True, wait=wait)
        engine = self._engine(handle)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle:
            self.assertTrue(orch.await_background_login(engine, {}, 5))
        settle.assert_called_once()

    def test_cancel_interrupts_wait(self):
        # 单次 handle.wait 最长可阻塞 150s 且不可取消——分片后取消立即生效
        from engine import cancel

        def wait(timeout=None):
            flag.write_text("cancel", encoding="utf-8")    # 等待中途收到取消请求
            return None

        handle = self._handle(finished=False, wait=wait)
        engine = self._engine(handle)
        with tempfile.TemporaryDirectory() as tmp:
            flag = Path(tmp) / "cancel.flag"
            cancel.set_cancel_file(flag)
            try:
                with self.assertRaises(cancel.CancelledError):
                    orch.await_background_login(engine, {}, 60)
            finally:
                cancel.set_cancel_file(None)

    def test_timeout_returns_false(self):
        handle = self._handle(finished=False, wait=lambda timeout=None: None)
        engine = self._engine(handle)
        with mock.patch.object(orch, "wait_ao_login_idle") as settle, \
             mock.patch.object(orch.time, "time",
                               side_effect=[100.0, 100.0 + 61]):
            # 第 1 次取算 deadline 的当前时间，第 2 次已超期 → 不再干等
            self.assertFalse(orch.await_background_login(engine, {}, 60))
        settle.assert_not_called()


if __name__ == "__main__":
    unittest.main()
