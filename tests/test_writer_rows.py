"""空白表单新增行 + 版本/锁定工具的纯逻辑测试（不碰 Excel/COM）。

覆盖用户提出的三条要求里的两类可离线验证的行为：
1. 编辑前先写维度列（组织/渠道/品牌/产品…），再写月度数据；
2. 版本名归一化（财总会版 == 财总版）与锁定提示解析。
"""
from __future__ import annotations

import re
import sys
import types
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError
from engine.writer import FormWriter
from engine.orchestrator import FormLockedError, _lock_user, _norm_version, _version_matches

WHITE = 0xFFFFFF
BLUE = 0xCC6600       # BGR：蓝分量 0xCC，深蓝只读底


class FakeRange:
    def __init__(self, sheet, addr: str) -> None:
        self.sheet = sheet
        self.addr = addr

    @property
    def Interior(self):
        color = self.sheet.colors.get(self.addr, self.sheet.color)
        return types.SimpleNamespace(Color=color)

    @property
    def Value2(self):
        if ":" not in self.addr:
            return self.sheet.cells.get(self.addr)
        start, end = self.addr.split(":")
        c1, r1 = _split_addr(start)
        c2, r2 = _split_addr(end)
        return tuple(
            tuple(self.sheet.cells.get(f"{_letter(c)}{r}") for c in range(c1, c2 + 1))
            for r in range(r1, r2 + 1)
        )

    @Value2.setter
    def Value2(self, v):
        self.sheet.cells[self.addr] = v


def _split_addr(addr: str) -> tuple[int, int]:
    """'AZ55' → (52, 55)。"""
    m = re.match(r"([A-Z]+)(\d+)", addr)
    return _col_index(m.group(1)), int(m.group(2))


def _col_index(addr: str) -> int:
    m = re.match(r"([A-Z]+)", addr)
    n = 0
    for ch in (m.group(1) if m else "A"):
        n = n * 26 + (ord(ch) - 64)
    return n


class FakeSheet:
    def __init__(self, cells: dict, color: int = WHITE,
                 colors: dict[str, int] | None = None) -> None:
        self.cells = dict(cells)
        self.color = color
        self.colors = dict(colors or {})
        used = [_col_index(a) for a in self.cells]
        first, last = (min(used), max(used)) if used else (1, 1)
        self.UsedRange = types.SimpleNamespace(
            Column=first,
            Columns=types.SimpleNamespace(Count=last - first + 1),
            Value2=None,
        )

    def Range(self, addr: str) -> FakeRange:
        return FakeRange(self, addr)

    def Cells(self, r: int, c: int) -> FakeRange:
        return FakeRange(self, f"{_letter(c)}{r}")

    def __getattr__(self, name):  # Shapes 等未被测试用到的属性
        if name == "Shapes":
            return []
        raise AttributeError(name)


class FakeEngine:
    def __init__(self, cells: dict | None = None, sheet: FakeSheet | None = None) -> None:
        self.sheet = sheet or FakeSheet(cells or {})
        self.write_calls: list[dict] = []
        self.saved = 0
        self.last_dialogs: list[str] = []
        self.macros_run: list[str] = []
        self.on_macro = None
        self._excel = types.SimpleNamespace(StatusBar="就绪")

    def run_macro(self, name, wb=None):
        self.macros_run.append(name)
        if self.on_macro:
            self.on_macro(name)

    def get_workbook(self, pattern, wait_sec=None):
        return types.SimpleNamespace(Name="fake.xlsx")

    def find_workbook(self, pattern):
        return types.SimpleNamespace(Name="fake.xlsx")

    def open_sheet(self, wb, name):
        return self.sheet

    def read(self, sheet, addr):
        return sheet.cells.get(addr)

    def write_batch(self, sheet, updates):
        self.write_calls.append(dict(updates))
        sheet.cells.update(updates)

    def save(self, wb, macro_hint=None):
        self.saved += 1

    def refresh(self, wb, macro_hint=None):
        pass


KEYED_CFG = {
    "workbook_pattern": "*",
    "sheet": "S",
    "layout": "keyed",
    "input_rows": [57, 62],
    "key_columns": [
        {"column": "AA", "name": "组织"},
        {"column": "AB", "name": "科目"},
    ],
    "month_columns": {"2027.01": "AI", "2027.02": "AJ"},
    "macros": {"save": "Plan_Save"},
}

SIMPLE_CFG = {
    "workbook_pattern": "*",
    "sheet": "S",
    "layout": "simple",
    "input_rows": [10, 12],
    "item_column": "P",
    "item_alternate_columns": ["Q"],
    "month_columns": {"2027.01": "AI"},
    "macros": {"save": "Plan_Save"},
}

# 注意：下面的 measures 故意写成旧的 6 列布局——动态识别应覆盖它。
MEASURE_CFG = {
    "workbook_pattern": "*",
    "sheet": "S",
    "layout": "measure_blocks",
    "header_row": 56,
    "measure_name_row": 55,
    "input_rows": [57, 60],
    "period_field": "期间",
    "key_columns": [
        {"column": "AA", "name": "组织"},
        {"column": "AB", "name": "渠道"},
        {"column": "AC", "name": "品牌"},
        {"column": "AD", "name": "产品"},
    ],
    "measures": {
        "销售单价": {"2026.11": "AH", "2026.12": "AI", "2027.01": "AJ",
                     "2027.02": "AK", "2027.03": "AL", "2027.04": "AM"},
        "销量": {"2026.11": "AN", "2026.12": "AO", "2027.01": "AP",
                 "2027.02": "AQ", "2027.03": "AR", "2027.04": "AS"},
        "加权CT": {"2026.11": "AT", "2026.12": "AU", "2027.01": "AV",
                   "2027.02": "AW", "2027.03": "AX", "2027.04": "AY"},
        "销售收入": {"2026.11": "AZ", "2026.12": "BA", "2027.01": "BB",
                     "2027.02": "BC", "2027.03": "BD", "2027.04": "BE"},
    },
    "measure_aliases": {
        "销售单价": ["销售单价（含税）"],
        "销量": ["销量(出货口径)"],
        "加权CT": ["加权CT"],
    },
    "macros": {"save": "Plan_Save"},
}

PERIODS_7 = ["2026.11", "2026.12", "2027.01", "2027.02", "2027.03", "2027.04", "2027.05"]


def _letter(n: int) -> str:
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _block(start: str, name: str, name_row: int = 55, header_row: int = 56) -> dict:
    first = _col_index(start)
    cells: dict[str, object] = {}
    for i, period in enumerate(PERIODS_7):
        col = _letter(first + i)
        cells[f"{col}{name_row}"] = name
        cells[f"{col}{header_row}"] = period
    return cells


class WriterNewRowTests(unittest.TestCase):
    def test_keyed_blank_form_allocates_rows_and_writes_dimensions_first(self) -> None:
        engine = FakeEngine()
        writer = FormWriter(engine, KEYED_CFG)
        data = [
            {"组织": "A", "科目": "X", "2027.01": 5},
            {"组织": "A", "科目": "Y", "2027.01": 7},
        ]
        pf = writer.preflight(data)
        self.assertTrue(pf.ok, pf.checks)
        self.assertEqual(len(pf.new_rows), 2)
        self.assertEqual(
            pf.key_updates,
            {"AA57": "A", "AB57": "X", "AA58": "A", "AB58": "Y"},
        )
        self.assertEqual(
            [(r["cell"], r["new_value"]) for r in pf.matched_rows],
            [("AI57", 5), ("AI58", 7)],
        )
        check = {name: (ok, detail) for name, ok, detail in pf.checks}
        self.assertIn("新增行", check)

        result = writer.write_and_save(pf)
        self.assertTrue(result["saved"])
        self.assertEqual(result["keys_written"], 4)
        # 第一批是维度，第二批才是月度数值
        self.assertEqual(engine.write_calls[0], pf.key_updates)
        self.assertEqual(engine.write_calls[1], {"AI57": 5, "AI58": 7})
        self.assertTrue(all(not addr.startswith(("AI", "AJ")) for addr in engine.write_calls[0]))

    def test_existing_key_reuses_row_and_skips_dimension_write(self) -> None:
        engine = FakeEngine({"AA57": "A", "AB57": "X", "AI57": 1})
        writer = FormWriter(engine, KEYED_CFG)
        pf = writer.preflight([{"组织": "A", "科目": "X", "2027.01": 5}])
        self.assertTrue(pf.ok, pf.checks)
        self.assertEqual(pf.key_updates, {})
        self.assertEqual(pf.new_rows, [])
        self.assertEqual(pf.matched_rows[0]["cell"], "AI57")
        self.assertEqual(pf.matched_rows[0]["old_value"], 1)

        result = writer.write_and_save(pf)
        self.assertNotIn("keys_written", result)
        self.assertEqual(engine.write_calls, [{"AI57": 5}])

    def test_missing_dimension_is_skipped_not_written(self) -> None:
        engine = FakeEngine()
        writer = FormWriter(engine, KEYED_CFG)
        pf = writer.preflight([{"组织": "A", "科目": "", "2027.01": 5}])
        self.assertFalse(pf.ok)
        self.assertEqual(pf.key_updates, {})
        self.assertIn("维度缺失", " ".join(pf.skipped))

    def test_no_blank_row_is_skipped(self) -> None:
        # 数据区只有一行且已被占用 → 新的维度组合明确报「无空白行」，不能串写到同一行
        engine = FakeEngine({"AA57": "A", "AB57": "X"})
        writer = FormWriter(engine, {**KEYED_CFG, "input_rows": [57, 57]})
        pf = writer.preflight([
            {"组织": "A", "科目": "X", "2027.01": 1},
            {"组织": "B", "科目": "Y", "2027.01": 2},
        ])
        self.assertFalse(pf.ok)
        self.assertIn("无空白行", " ".join(pf.skipped))
        self.assertEqual(len(pf.new_rows), 0)

    def test_simple_form_blank_check_uses_item_column_only(self) -> None:
        # AA 有分组标题不算占行：只要科目列 P 为空，该行就可新增
        engine = FakeEngine({"AA10": "分组标题"})
        writer = FormWriter(engine, SIMPLE_CFG)
        pf = writer.preflight([{"item": "维修费", "2027.01": 3}])
        self.assertTrue(pf.ok, pf.checks)
        self.assertEqual(pf.key_updates, {"P10": "维修费"})
        self.assertEqual(pf.new_rows[0]["row"], 10)

    def test_second_preflight_resets_new_row_state(self) -> None:
        engine = FakeEngine()
        writer = FormWriter(engine, KEYED_CFG)
        first = writer.preflight([{"组织": "A", "科目": "X", "2027.01": 1}])
        second = writer.preflight([{"组织": "B", "科目": "Y", "2027.01": 2}])
        self.assertEqual(list(first.key_updates), ["AA57", "AB57"])
        self.assertEqual(list(second.key_updates), ["AA57", "AB57"])
        self.assertEqual(len(second.new_rows), 1)


class MeasureBlockHeaderTests(unittest.TestCase):
    """价格&销量表的度量块宽度随版本变化：必须按表头动态识别列。"""

    ROW = {
        "组织": "A", "渠道": "B", "品牌": "C", "产品": "P", "期间": "2027.02",
        "销售单价（含税）": 100, "销量(出货口径)": 5, "加权CT": 80,
    }

    def _seven_col_headers(self) -> dict:
        cells: dict = {}
        cells.update(_block("AH", "销售单价（含税）"))
        cells.update(_block("AO", "销量(出货口径)\n（台）"))
        cells.update(_block("AV", "加权CT"))
        cells.update(_block("BC", "销售收入"))
        return cells

    def test_header_detection_maps_7_column_blocks(self) -> None:
        engine = FakeEngine(self._seven_col_headers())
        writer = FormWriter(engine, MEASURE_CFG)
        pf = writer.preflight([self.ROW])
        self.assertTrue(pf.ok, pf.checks)
        self.assertEqual(
            {r["cell"] for r in pf.matched_rows},
            {"AK57", "AR57", "AY57"},          # 2027.02：单价/销量/加权CT
        )

    def test_unreadable_headers_raise_instead_of_fallback(self) -> None:
        # 表头读不到（表单未加载/行号配错）时直接报错，绝不回退配置列：
        # 模板块宽随版本期间窗口变（6↔7 列），配置列可能已与模板错位，
        # 硬写配置列会把数字写进错误的度量/期间——宁可中止也不能写错。
        engine = FakeEngine()
        writer = FormWriter(engine, MEASURE_CFG)
        with self.assertRaises(ExcelError) as ctx:
            writer.preflight([self.ROW])
        self.assertIn("度量块表头识别失败", str(ctx.exception))
        self.assertEqual(engine.write_calls, [])   # 一个字都不能写

    def test_partial_headers_raise_instead_of_fallback(self) -> None:
        # 只有部分度量块有表头（销售收入块缺失）→ 报错，不能拿配置列补写
        cells = self._seven_col_headers()
        for col in ("BC", "BD", "BE", "BF", "BG", "BH", "BI"):
            cells.pop(f"{col}55", None)
            cells.pop(f"{col}56", None)
        engine = FakeEngine(cells)
        writer = FormWriter(engine, MEASURE_CFG)
        with self.assertRaises(ExcelError) as ctx:
            writer.preflight([self.ROW])
        self.assertIn("销售收入", str(ctx.exception))
        self.assertEqual(engine.write_calls, [])

    def test_inconsistent_block_widths_raise(self) -> None:
        # 某个度量块的期间数与其它块不一致（合并单元格/空列的典型症状）→ 报错
        cells = self._seven_col_headers()
        cells.pop("BI55")      # 销售收入块少一个期间（只剩 6 个）
        cells.pop("BI56")
        engine = FakeEngine(cells)
        writer = FormWriter(engine, MEASURE_CFG)
        with self.assertRaises(ExcelError) as ctx:
            writer.preflight([self.ROW])
        self.assertIn("期间数不一致", str(ctx.exception))

    def test_detected_periods_are_used_for_coverage_check(self) -> None:
        engine = FakeEngine(self._seven_col_headers())
        writer = FormWriter(engine, MEASURE_CFG)
        pf = writer.preflight([{**self.ROW, "期间": "2027.05"}])
        self.assertTrue(pf.ok, pf.checks)
        self.assertEqual(
            {r["cell"] for r in pf.matched_rows},
            {"AN57", "AU57", "BB57"},          # 表头识别出的 2027.05 列
        )
        self.assertNotIn("配置缺少期间", " ".join(d for _, _, d in pf.checks))


class VersionAndLockTests(unittest.TestCase):
    def test_version_name_normalisation(self) -> None:
        self.assertEqual(
            _norm_version("滚动预测 2026年11月 财总会版"),
            _norm_version("滚动预测2026年11月财总版"),
        )
        self.assertTrue(_version_matches("滚动预测2026年11月财总会版", "滚动预测2026年11月财总版"))
        self.assertTrue(_version_matches("滚动预测2026年11月财总版", "滚动预测2026年11月财总版"))
        self.assertFalse(_version_matches("滚动预测2026年11月财总版", "滚动预测2026年12月财总版"))
        self.assertFalse(_version_matches("", "滚动预测2026年11月财总版"))
        self.assertFalse(_version_matches("滚动预测2026年11月财总版", ""))

    def test_lock_user_extraction(self) -> None:
        self.assertEqual(_lock_user("表单被张三用户锁定，请稍后再试"), "张三")
        self.assertEqual(_lock_user("该表单被 Zhang San 锁定"), "Zhang San")
        self.assertEqual(_lock_user("Form is locked by zhangsan"), "zhangsan")
        self.assertEqual(_lock_user("没有相关信息"), "")

    def test_locked_error_message_mentions_form_and_user(self) -> None:
        exc = FormLockedError(
            "费用及损益其他项编制_非海外",
            "变动费用",
            "表单被张三用户锁定，请稍后再试",
        )
        self.assertEqual(exc.user, "张三")
        self.assertIn("变动费用表单被张三用户锁定", str(exc))
        self.assertIn("结束流程", str(exc))
        self.assertIn("原始提示", str(exc))

    def test_lock_detection_patterns(self) -> None:
        looks = FormWriter._looks_lock
        self.assertTrue(looks("该表单已被张三用户锁定"))
        self.assertTrue(looks("This form is locked by zhangsan"))
        self.assertTrue(looks("表单正在编辑中"))
        self.assertFalse(looks("表单已刷新"))
        self.assertFalse(looks(""))
        self.assertFalse(looks(None))


MODE_CFG = {
    "workbook_pattern": "*",
    "sheet": "S",
    "mode_display_cell": "AM51",
    "mode_switch": {"macro": "Switch_mode"},
    "macros": {"switch_mode": "Switch_mode", "save": "Plan_Save"},
    "mode_poll_interval_sec": 0,
    "mode_poll_attempts": 2,
}


class ModeSwitchTests(unittest.TestCase):
    """查询/编制模式判定：价格表查询态蓝色输入区必须切换后才能保存。

    2026-11-23 实测：价格&销量&加权CT 表在“查询”模式下，单价列白底、
    销量/CT 列蓝底；只看第一个目标格会误判为可编辑，Plan_Save 弹
    「没有要保存的数据」。客户费用表则相反——查询模式也能直接写入并保存。
    """

    def _engine(self, mode: str = "查询", colors: dict | None = None) -> FakeEngine:
        sheet = FakeSheet({"AL51": "当前模式:", "AM51": mode}, colors=colors)
        return FakeEngine(sheet=sheet)

    def test_query_with_blue_cells_runs_switch_macro(self) -> None:
        engine = self._engine(colors={"AO57": BLUE, "AX57": BLUE})
        engine.on_macro = lambda name: engine.sheet.cells.__setitem__("AM51", "编制")
        writer = FormWriter(engine, MODE_CFG)
        writer.enter_edit_mode(sample_cells=["AJ57", "AO57", "AX57"])
        self.assertEqual(engine.macros_run, ["Switch_mode"])

    def test_edit_mode_skips_switch_even_if_cells_blue(self) -> None:
        # 编制态下再调 Switch_mode 会反向切回查询，必须跳过
        engine = self._engine("编制", colors={"AO57": BLUE})
        writer = FormWriter(engine, MODE_CFG)
        writer.enter_edit_mode(sample_cells=["AO57"])
        self.assertEqual(engine.macros_run, [])

    def test_query_with_white_cells_skips_switch(self) -> None:
        # 客户费用表：查询态输入区仍为白底，AO 允许直接写入并保存
        engine = self._engine()
        writer = FormWriter(engine, MODE_CFG)
        writer.enter_edit_mode(sample_cells=["AJ57", "AI57"])
        self.assertEqual(engine.macros_run, [])

    def test_query_blue_without_macro_raises(self) -> None:
        cfg = {**MODE_CFG, "mode_switch": {"macro": None}, "macros": {}}
        engine = self._engine(colors={"AO57": BLUE})
        writer = FormWriter(engine, cfg)
        with self.assertRaises(ExcelError):
            writer.enter_edit_mode(sample_cells=["AO57"])

    def test_switch_ineffective_raises(self) -> None:
        engine = self._engine(colors={"AO57": BLUE})
        writer = FormWriter(engine, MODE_CFG)   # 宏运行后 AM51 仍是“查询”
        with self.assertRaises(ExcelError) as ctx:
            writer.enter_edit_mode(sample_cells=["AO57"])
        self.assertIn("未能进入编制态", str(ctx.exception))

    def test_auto_scan_finds_mode_display(self) -> None:
        # 不配 mode_display_cell 时，扫描表头上方文本里的「当前模式:」
        cfg = {k: v for k, v in MODE_CFG.items() if k != "mode_display_cell"}
        cfg["header_row"] = 56
        engine = self._engine("编制", colors={"AO57": BLUE})
        writer = FormWriter(engine, cfg)
        writer.enter_edit_mode(sample_cells=["AO57"])
        self.assertEqual(engine.macros_run, [])


if __name__ == "__main__":
    unittest.main()
