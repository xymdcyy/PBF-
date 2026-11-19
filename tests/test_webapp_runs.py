"""Web 运行注册表：并发抢占、释放、淘汰。"""
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from webapp.runs import RunRegistry


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
