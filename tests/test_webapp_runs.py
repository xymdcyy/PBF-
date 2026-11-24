"""Web 运行注册表：并发抢占、释放、淘汰。"""
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from webapp.runs import (RunRegistry, can_force_release, run_owned_pids,
                         should_hold_slot)


class RunOwnedPidsTest(unittest.TestCase):
    def test_only_new_pids_are_owned(self):
        # 运行前就在的 Excel（用户自己开的）不算本次运行所有
        state = {"excel_pids_before": [100, 200]}
        self.assertEqual(run_owned_pids(state, [100, 200, 300]), {300})

    def test_baseline_gone_owns_nothing(self):
        state = {"excel_pids_before": [100]}
        self.assertEqual(run_owned_pids(state, [100]), set())

    def test_persisted_owned_pids_counted(self):
        # 归属记录里已登记的 pbf 实例也计入
        state = {"excel_pids_before": [100], "excel_owned_pids": [300]}
        self.assertEqual(run_owned_pids(state, [100, 300]), {300})

    def test_dead_owned_pid_not_returned(self):
        state = {"excel_pids_before": [100], "excel_owned_pids": [300]}
        self.assertEqual(run_owned_pids(state, [100]), set())

    def test_unknown_baseline_is_conservative(self):
        # 运行开始时枚举进程失败：无法区分用户 Excel，宁可多占槽位
        state = {"excel_pids_before": None}
        self.assertEqual(run_owned_pids(state, [100]), {100})

    def test_none_current_means_unknown(self):
        self.assertEqual(run_owned_pids({"excel_pids_before": [100]}, None), set())


class ShouldHoldSlotTest(unittest.TestCase):
    def test_alive_proc_holds(self):
        # 子进程没被确认终止：无论如何都不放行，避免两个流程操作同一个 Excel
        self.assertTrue(should_hold_slot({}, proc_alive=True, current_pids=[]))
        self.assertTrue(should_hold_slot({}, proc_alive=True, current_pids=None))

    def test_cancelled_holds_while_run_owned_excel_alive(self):
        # Excel 由 COM 启动、不在 taskkill /T 的进程树内：本次运行的 Excel
        # 没退出前占位
        state = {"cancelled": True, "excel_pids_before": [100]}
        self.assertTrue(should_hold_slot(state, proc_alive=False,
                                         current_pids=[100, 300]))
        self.assertFalse(should_hold_slot(state, proc_alive=False,
                                          current_pids=[100]))

    def test_cancelled_unknown_pids_holds(self):
        self.assertTrue(should_hold_slot({"cancelled": True}, proc_alive=False,
                                         current_pids=None))

    def test_cancelled_graceful_releases_even_with_excel_alive(self):
        # 子进程已在安全点自行退出，Excel 可能还开着导航表，下一轮复用即可
        state = {"cancelled": True, "cancel_graceful": True,
                 "excel_pids_before": []}
        self.assertFalse(should_hold_slot(state, proc_alive=False,
                                          current_pids=[300]))

    def test_cancelled_package_running_holds(self):
        # 运行包已在 BPC 后台触发，本地探不到进度，保守占位
        state = {"cancelled": True, "cancel_package_running": True,
                 "excel_pids_before": [100]}
        self.assertTrue(should_hold_slot(state, proc_alive=False,
                                         current_pids=[100]))

    def test_timed_out_holds_until_run_owned_excel_exits(self):
        state = {"timed_out": True, "excel_pids_before": [100]}
        self.assertTrue(should_hold_slot(state, proc_alive=False,
                                         current_pids=[100, 300]))
        self.assertFalse(should_hold_slot(state, proc_alive=False,
                                          current_pids=[100]))

    def test_normal_completion_releases(self):
        self.assertFalse(should_hold_slot({}, proc_alive=False,
                                          current_pids=[300]))


class CanForceReleaseTest(unittest.TestCase):
    def test_stuck_finished_dead_proc_can_release(self):
        self.assertTrue(can_force_release({"done": True, "stuck": True},
                                          proc_alive=False))

    def test_live_proc_cannot_release(self):
        self.assertFalse(can_force_release({"done": True, "stuck": True},
                                           proc_alive=True))

    def test_not_stuck_cannot_release(self):
        self.assertFalse(can_force_release({"done": True, "stuck": False},
                                           proc_alive=False))

    def test_not_done_cannot_release(self):
        self.assertFalse(can_force_release({"done": False, "stuck": True},
                                           proc_alive=False))


class RunRegistryTest(unittest.TestCase):
    def test_begin_is_exclusive(self):
        reg = RunRegistry()
        self.assertTrue(reg.begin("a", {"done": False}))
        self.assertFalse(reg.begin("b", {"done": False}))
        self.assertEqual(reg.current, "a")
        self.assertIsNone(reg.get("b"))

    def test_end_releases_slot(self):
        reg = RunRegistry()
        reg.begin("a", {"done": False})
        reg.end("a")
        self.assertIsNone(reg.current)
        self.assertTrue(reg.begin("b", {"done": False}))
        self.assertEqual(reg.current, "b")

    def test_end_only_releases_matching_run(self):
        reg = RunRegistry()
        reg.begin("a", {"done": False})
        reg.end("other")
        self.assertEqual(reg.current, "a")

    def test_prune_keeps_recent_finished_runs(self):
        reg = RunRegistry(max_kept=3)
        for i in range(5):
            rid = f"r{i}"
            reg.begin(rid, {"done": True})
            reg.end(rid)
        # 旧的已结束运行被淘汰，只保留最近 3 个
        self.assertIsNone(reg.get("r0"))
        self.assertIsNone(reg.get("r1"))
        self.assertIsNotNone(reg.get("r4"))

    def test_prune_never_drops_current_run(self):
        reg = RunRegistry(max_kept=1)
        reg.begin("running", {"done": False})
        reg.begin("done", {"done": True})  # 抢占失败，不注册
        self.assertIsNotNone(reg.get("running"))
        self.assertIsNone(reg.get("done"))


if __name__ == "__main__":
    unittest.main()
