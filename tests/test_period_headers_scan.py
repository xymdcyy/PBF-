"""_read_period_headers 的健壮性：异常表头跳过、非期间键丢弃。

- 转换类异常（巨大整数 float() 溢出等）只跳过该格，不中断整轮核对；
- NaN → "nan"、杂散文本等非期间键不得混进表头表（会污染报错信息）。
"""
from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.orchestrator import PbfFlow

HEADER_ROW = 54


class _FakeWS:
    pass


def _make_flow(values: dict[str, object]) -> PbfFlow:
    """只支持 _read 的最小替身：按 'A54' 这类地址返回预设值，缺省 None。"""
    flow = object.__new__(PbfFlow)
    flow._read = lambda ws, addr: values.get(addr)
    return flow


class PeriodHeadersScanTests(unittest.TestCase):
    def test_keeps_period_keys_and_skips_garbage(self) -> None:
        values = {
            "A54": "2027.03",              # 文本期间
            "B54": 2027.04,                # 浮点期间
            "C54": datetime(2027, 5, 1),   # 日期期间
            "G54": 2026,                   # 纯年份
            "D54": float("nan"),           # 异常单元格 → "nan"，应被丢弃
            "E54": float("inf"),           # → "inf"，应被丢弃
            "F54": "指标",                  # 杂散文本，应被丢弃
        }
        found = _make_flow(values)._read_period_headers(_FakeWS(), HEADER_ROW)
        self.assertEqual(found, {
            "2027.03": "A", "2027.04": "B", "2027.05": "C", "2026": "G",
        })

    def test_conversion_error_skips_cell(self) -> None:
        # 巨大整数在 normalize 里 float() 会 OverflowError：只跳过该格
        values = {"A54": "2027.03", "B54": 10 ** 400}
        found = _make_flow(values)._read_period_headers(_FakeWS(), HEADER_ROW)
        self.assertEqual(found, {"2027.03": "A"})


if __name__ == "__main__":
    unittest.main()
