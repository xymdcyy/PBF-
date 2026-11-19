"""批量清除：拒绝覆盖源目录，并清理输出目录里的陈旧 CSV（P1-8 配套）。"""
from __future__ import annotations

import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.clear import FORM_SCHEMA, clear_data


class ClearTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        base = Path(self._tmp.name)
        self.src = base / "data_2027.01"
        self.src.mkdir()
        with open(self.src / "客户费用.csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["渠道", "科目", "2027.01"])
            w.writeheader()
            w.writerow({"渠道": "电商", "科目": "返利", "2027.01": 12})
        self.out = base / "data_2027.01_clear"
        self.out.mkdir()

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_refuses_output_equal_to_source(self) -> None:
        with self.assertRaises(ValueError):
            clear_data([self.src], ["2027.01"], self.src)
        # 原始数据必须保持原样
        self.assertIn("12", (self.src / "客户费用.csv").read_text(encoding="utf-8-sig"))

    def test_raises_when_no_keys_found(self) -> None:
        empty = Path(self._tmp.name) / "data_2026.12"
        empty.mkdir()
        with self.assertRaises(ValueError):
            clear_data([empty], ["2027.01"], self.out)

    def test_removes_stale_files_and_writes_zeros(self) -> None:
        for name in FORM_SCHEMA:
            (self.out / name).write_text("OLD\n", encoding="utf-8")
        rows = clear_data([self.src], ["2027.01"], self.out)
        self.assertEqual(rows.get("客户费用.csv"), 1)
        for name in FORM_SCHEMA:
            self.assertFalse((self.out / name).exists() and name not in rows,
                             f"{name} 没有生成却残留旧文件")
        text = (self.out / "客户费用.csv").read_text(encoding="utf-8-sig")
        self.assertNotIn("OLD", text)
        self.assertIn(",0", text.replace("﻿", ""))


if __name__ == "__main__":
    unittest.main()
