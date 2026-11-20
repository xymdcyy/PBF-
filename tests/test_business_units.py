"""业务单元（品牌 × 品类）配置：覆盖、校验与导出空列告警。"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(name: str, filename: str):
    path = ROOT / "engine" / filename
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


bu = _load("business_units_under_test", "business_units.py")


def _flow() -> dict:
    return {
        "steps": [
            {"form": "E61_CBG_价格销量CT编制", "nav_cell": "AC55"},
            {"form": "HQ_毛上调整", "nav_cell": "AC59",
             "after_close": {"nav_cell": "AC62", "packages": [52, 53]}},
            {"form": "费用及损益其他项编制_非海外", "nav_cell": "AC65",
             "after_close": {"nav_cell": "AC70", "packages": [52]}},
        ],
        "verify": {"nav_cell": "AC70", "form_row": 54},
    }


def _units() -> dict:
    return {
        "品牌": ["TCL", "雷鸟", "乐华"],
        "品类": ["电视", "空调"],
        "units": {
            "乐华·电视": {
                "启用": True, "品牌": "乐华", "品类": "电视",
                "verify_nav_cell": "B70", "verify_form_row": 55,
                "steps": [
                    {"表单": "价格销量CT", "nav_cell": "B55"},
                    {"表单": "HQ毛上调整", "nav_cell": "B59", "after_close_nav_cell": "B62"},
                    {"表单": "费用及损益", "nav_cell": "B65", "after_close_nav_cell": "B70"},
                ],
            },
            "TCL·电视": {"启用": False, "品牌": "TCL", "品类": "电视"},
            "TCL·空调": {"启用": False, "品牌": "TCL", "品类": "空调"},
        },
    }


class TestLookup(unittest.TestCase):
    def test_enabled_and_default(self):
        cfg = _units()
        self.assertEqual(bu.enabled_unit_names(cfg), ["乐华·电视"])
        self.assertEqual(bu.default_unit(cfg), "乐华·电视")
        self.assertEqual(bu.unit_names(cfg),
                         ["乐华·电视", "TCL·电视", "TCL·空调"])

    def test_no_enabled_unit(self):
        cfg = _units()
        for u in cfg["units"].values():
            u["启用"] = False
        self.assertIsNone(bu.default_unit(cfg))

    def test_get_unknown_unit(self):
        with self.assertRaises(bu.UnitError):
            bu.get_unit(_units(), "雷鸟·电视")

    def test_get_disabled_unit(self):
        with self.assertRaises(bu.UnitError):
            bu.get_unit(_units(), "TCL·电视")


class TestApplyUnit(unittest.TestCase):
    def test_overrides_nav_cells(self):
        flow = bu.apply_unit(_flow(), _units(), "乐华·电视")
        self.assertEqual([s["nav_cell"] for s in flow["steps"]], ["B55", "B59", "B65"])
        self.assertEqual(flow["steps"][1]["after_close"]["nav_cell"], "B62")
        self.assertEqual(flow["steps"][2]["after_close"]["nav_cell"], "B70")
        # packages 不被单元配置破坏
        self.assertEqual(flow["steps"][1]["after_close"]["packages"], [52, 53])
        self.assertEqual(flow["verify"]["nav_cell"], "B70")
        self.assertEqual(flow["verify"]["form_row"], 55)
        self.assertEqual(flow["_unit"],
                         {"name": "乐华·电视", "brand": "乐华", "category": "电视"})

    def test_does_not_mutate_inputs(self):
        base = _flow()
        bu.apply_unit(base, _units(), "乐华·电视")
        self.assertEqual(base["steps"][0]["nav_cell"], "AC55")
        self.assertEqual(base["verify"]["nav_cell"], "AC70")
        self.assertNotIn("_unit", base)

    def test_without_optional_overrides(self):
        units = _units()
        unit = units["units"]["乐华·电视"]
        unit.pop("verify_nav_cell")
        unit.pop("verify_form_row")
        for step in unit["steps"]:
            step.pop("after_close_nav_cell", None)
        flow = bu.apply_unit(_flow(), units, "乐华·电视")
        self.assertEqual(flow["steps"][1]["after_close"]["nav_cell"], "AC62")
        self.assertEqual(flow["verify"]["nav_cell"], "AC70")

    def test_step_count_mismatch(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"].pop()
        with self.assertRaisesRegex(bu.UnitError, "steps 数量"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_step_form_mismatch(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][1]["表单"] = "客户费用"
        with self.assertRaisesRegex(bu.UnitError, "不对应"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_missing_nav_cell(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"][0]["nav_cell"] = ""
        with self.assertRaisesRegex(bu.UnitError, "nav_cell"):
            bu.apply_unit(_flow(), units, "乐华·电视")

    def test_disabled_unit_rejected(self):
        with self.assertRaises(bu.UnitError):
            bu.apply_unit(_flow(), _units(), "TCL·电视")


class TestCheckUnits(unittest.TestCase):
    def test_valid_config_passes(self):
        self.assertEqual(bu.check_units(_units(), _flow()), [])

    def test_missing_brand_or_category(self):
        units = _units()
        units["units"]["TCL·电视"].pop("品类")
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("缺少 品牌/品类" in i for i in issues))

    def test_brand_not_in_list(self):
        units = _units()
        units["units"]["TCL·空调"]["品牌"] = "索尼"
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("不在 品牌 列表中" in i for i in issues))

    def test_enabled_unit_with_bad_steps(self):
        units = _units()
        units["units"]["乐华·电视"]["steps"] = units["units"]["乐华·电视"]["steps"][:1]
        issues = bu.check_units(units, _flow())
        self.assertTrue(any(i.startswith("错误:") and "steps 数量" in i for i in issues))

    def test_duplicate_nav_cell(self):
        units = _units()
        steps = units["units"]["乐华·电视"]["steps"]
        steps[1]["nav_cell"] = steps[0]["nav_cell"]
        issues = bu.check_units(units, _flow())
        self.assertTrue(any("nav_cell 重复" in i for i in issues))

    def test_no_units(self):
        issues = bu.check_units({}, _flow())
        self.assertTrue(any("未配置任何业务单元" in i for i in issues))


class TestCheckUnitsOnly(unittest.TestCase):
    """按 --unit 运行时只校验被选中的单元，其余单元的问题不阻塞。"""

    def test_only_ignores_other_units(self):
        units = _units()
        # 未启用的占位单元缺品类：全量校验报错，only 校验不应报
        units["units"]["TCL·电视"].pop("品类")
        self.assertTrue(any("缺少 品牌/品类" in i
                            for i in bu.check_units(units, _flow())))
        self.assertEqual(bu.check_units(units, _flow(), only="乐华·电视"), [])

    def test_only_reports_selected_unit_issues(self):
        units = _units()
        steps = units["units"]["乐华·电视"]["steps"]
        steps[1]["nav_cell"] = steps[0]["nav_cell"]
        issues = bu.check_units(units, _flow(), only="乐华·电视")
        self.assertTrue(any("nav_cell 重复" in i for i in issues))

    def test_only_reports_selected_unit_missing_dims(self):
        units = _units()
        units["units"]["乐华·电视"].pop("品类")
        issues = bu.check_units(units, _flow(), only="乐华·电视")
        self.assertTrue(any("缺少 品牌/品类" in i for i in issues))

    def test_only_unknown_unit(self):
        issues = bu.check_units(_units(), _flow(), only="雷鸟·电视")
        self.assertTrue(any("未知业务单元" in i for i in issues))

    def test_only_skips_global_no_enabled_check(self):
        # only 指定的单元本身是启用的，全局“没有启用单元”检查不参与
        units = _units()
        for name, u in units["units"].items():
            if name != "乐华·电视":
                u["启用"] = False
        self.assertEqual(bu.check_units(units, _flow(), only="乐华·电视"), [])


class TestCsvDimensionWarnings(unittest.TestCase):
    def _write(self, directory: Path, name: str, text: str) -> None:
        (directory / name).write_text(text, encoding="utf-8-sig")

    def test_empty_brand_and_category(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "价格销量CT.csv",
                        "组织,渠道,品牌,产品,期间,销量\n组织A,渠道A,,产品A,2027.01,10\n")
            warnings = bu.csv_dimension_warnings(d)
            self.assertEqual(len(warnings), 1)
            self.assertIn("品牌", warnings[0])
            self.assertIn("全部为空", warnings[0])

    def test_filled_dimensions_no_warning(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "a.csv",
                        "品牌,品类,期间,值\n乐华,电视,2027.01,1\n")
            self.assertEqual(bu.csv_dimension_warnings(d), [])

    def test_missing_column_no_warning(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            self._write(d, "b.csv", "科目,期间,值\nA,2027.01,1\n")
            self.assertEqual(bu.csv_dimension_warnings(d), [])

    def test_gbk_csv(self):
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            (d / "c.csv").write_bytes("品牌,期间\n,2027.01\n".encode("gbk"))
            warnings = bu.csv_dimension_warnings(d)
            self.assertEqual(len(warnings), 1)


if __name__ == "__main__":
    unittest.main()
