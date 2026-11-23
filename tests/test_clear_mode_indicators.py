"""清 0 模式的损益核对放行规则。

背景（2026-11-23 现场）：网页 clear 把 2027.01-02 五张表填 0 后，流程最后
还要跑「损益核对」。核对里有一条"目标期间指标全 0 说明运行包没生效"的
失败判定——清 0 场景下全 0 正是预期结果，不放行会让一次成功的清 0 被报成失败。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError
from engine.orchestrator import PbfFlow


def make_flow() -> PbfFlow:
    cfg = {
        "navigator": {
            "workbook_pattern": "PBF Pro*导航*",
            "sheet": "nav",
            "open_button_cell": "AT52",
            "refresh_flow_macro": "Refresh_Status",
            "open_wait_sec": 1,
            "version_cell": "AR45",
        },
        "package": {"workbook_pattern": "PBF Pro运行包*", "sheet": "pkg"},
        "verify": {"indicator_rows": {"净收入": 69, "端到端净毛利": 86}},
    }
    return PbfFlow(object(), cfg, {}, Path("data_2027.01_2027.02_clear"))


ALL_ZERO = {"净收入": {"2027.02": 0.0}, "端到端净毛利": {"2027.02": None}}


class ClearModeIndicatorTests(unittest.TestCase):
    def test_all_zero_rejected_by_default(self) -> None:
        flow = make_flow()
        with self.assertRaises(ExcelError) as ctx:
            flow._check_target_indicators(ALL_ZERO, "2027.02", allow_all_zero=False)
        self.assertIn("全部为空或为 0", str(ctx.exception))
        self.assertFalse(flow.indicators_all_zero)

    def test_all_zero_allowed_in_clear_mode(self) -> None:
        flow = make_flow()
        self.assertTrue(flow._check_target_indicators(
            ALL_ZERO, "2027.02", allow_all_zero=True))
        self.assertTrue(flow.indicators_all_zero)

    def test_partial_zero_stays_non_clear(self) -> None:
        flow = make_flow()
        out = {"净收入": {"2027.02": 0.0}, "端到端净毛利": {"2027.02": 123.4}}
        self.assertFalse(flow._check_target_indicators(
            out, "2027.02", allow_all_zero=True))
        self.assertFalse(flow.indicators_all_zero)


if __name__ == "__main__":
    unittest.main()
