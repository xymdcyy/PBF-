"""离线配置/数据校验。"""
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / "engine" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    # dataclass 在处理注解时需要能从 sys.modules 找到本模块
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


cc = _load("config_check")


def levels(issues):
    return [i.level for i in issues]


def messages(issues):
    return " | ".join(i.message for i in issues)


KEYED = {
    "form_id": "F1",
    "layout": "keyed",
    "workbook_pattern": "F1*",
    "sheet": "S1",
    "nav_cell": "AC57",
    "input_rows": [58, 80],
    "key_columns": [{"column": "AA", "name": "组织"}, {"column": "AB", "name": "科目"}],
    "month_columns": {"2027.03": "AD", "2027.04": "AE"},
}

MEASURES = {
    "form_id": "F2",
    "layout": "measure_blocks",
    "workbook_pattern": "F2*",
    "sheet": "S2",
    "nav_cell": "AC55",
    "input_rows": [3, 50],
    "key_columns": [{"column": "AA", "name": "组织"}, {"column": "AB", "name": "产品"}],
    "measures": {"单价": {"2027.03": "AD"}, "销量": {"2027.03": "AE"}},
}


class PeriodsTest(unittest.TestCase):
    def test_configured_periods_keyed(self):
        self.assertEqual(cc.configured_periods(KEYED), {"2027.03", "2027.04"})

    def test_configured_periods_measure_blocks(self):
        self.assertEqual(cc.configured_periods(MEASURES), {"2027.03"})

    def test_configured_periods_matrix(self):
        cfg = {"layout": "matrix", "period_blocks": {"2027": "AW", "2027.03": "BR"}}
        self.assertEqual(cc.configured_periods(cfg), {"2027", "2027.03"})

    def test_data_periods_wide(self):
        rows = [{"组织": "a", "2027.03": 1, "2027.04": 2}]
        self.assertEqual(cc.data_periods(rows, KEYED), {"2027.03", "2027.04"})

    def test_data_periods_long(self):
        rows = [{"组织": "a", "期间": "2027.03"}, {"组织": "b", "期间": "2027.04"}]
        self.assertEqual(cc.data_periods(rows, MEASURES), {"2027.03", "2027.04"})


class FormConfigTest(unittest.TestCase):
    def test_valid_configs_have_no_errors(self):
        self.assertEqual(levels(cc.check_form_config("F1", KEYED)), [])
        self.assertEqual(levels(cc.check_form_config("F2", MEASURES)), [])

    def test_missing_required_keys(self):
        issues = cc.check_form_config("bad", {"layout": "simple"})
        self.assertIn("error", levels(issues))
        msg = messages(issues)
        for key in ("workbook_pattern", "sheet", "nav_cell", "item_column"):
            self.assertIn(key, msg)

    def test_bad_layout(self):
        issues = cc.check_form_config("bad", dict(KEYED, layout="unknown"))
        self.assertIn("error", levels(issues))

    def test_duplicate_period_column(self):
        cfg = dict(KEYED, month_columns={"2027.03": "AD", "2027.04": "AD"})
        issues = cc.check_form_config("F1", cfg)
        self.assertIn("列号重复", messages(issues))


class DataFileTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def _csv(self, name: str, text: str) -> Path:
        path = self.dir / name
        path.write_text(text, encoding="utf-8-sig")
        return path

    def test_ok(self):
        path = self._csv("a.csv", "组织,科目,2027.03\nA,B,1\n")
        issues, rows = cc.check_data_file(path, KEYED)
        self.assertEqual(levels(issues), [])
        self.assertEqual(len(rows), 1)

    def test_missing_columns(self):
        path = self._csv("a.csv", "组织,2027.03\nA,1\n")
        issues, _ = cc.check_data_file(path, KEYED)
        self.assertIn("缺少列", messages(issues))

    def test_uncovered_period(self):
        path = self._csv("a.csv", "组织,科目,2027.05\nA,B,1\n")
        issues, _ = cc.check_data_file(path, KEYED)
        self.assertIn("2027.05", messages(issues))
        self.assertIn("没有对应列", messages(issues))

    def test_duplicate_key(self):
        path = self._csv("a.csv", "组织,科目,2027.03\nA,B,1\nA,B,2\n")
        issues, _ = cc.check_data_file(path, KEYED)
        self.assertIn("重复维度键", messages(issues))

    def test_missing_file(self):
        issues, rows = cc.check_data_file(self.dir / "nope.csv", KEYED)
        self.assertIn("不存在", messages(issues))
        self.assertEqual(rows, [])

    def test_empty_file(self):
        path = self._csv("a.csv", "组织,科目,2027.03\n")
        issues, _ = cc.check_data_file(path, KEYED)
        self.assertIn("为空", messages(issues))

    def test_target_period_warning(self):
        path = self._csv("a.csv", "组织,科目,2027.03\nA,B,1\n")
        issues, _ = cc.check_data_file(path, KEYED, target_period="2027.04")
        self.assertEqual(levels(issues), ["warning"])

    def test_measure_long_duplicate(self):
        path = self._csv(
            "m.csv", "组织,产品,期间,单价,销量\nA,X,2027.03,1,2\nA,X,2027.03,3,4\n")
        issues, _ = cc.check_data_file(path, MEASURES)
        self.assertIn("重复维度键", messages(issues))

    def test_measure_missing_measure_column(self):
        path = self._csv("m.csv", "组织,产品,期间,单价\nA,X,2027.03,1\n")
        issues, _ = cc.check_data_file(path, MEASURES)
        self.assertIn("销量", messages(issues))

    def test_measure_alias_column_accepted(self):
        cfg = dict(MEASURES, measure_aliases={"销量": ["销量(出货口径)"]})
        path = self._csv(
            "m.csv", "组织,产品,期间,单价,销量(出货口径)\nA,X,2027.03,1,2\n")
        issues, _ = cc.check_data_file(path, cfg)
        self.assertEqual(levels(issues), [], messages(issues))


class FlowTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.nav = self.root / "nav.xlsb"
        self.nav.write_bytes(b"")
        self.data_dir = self.root / "data_2027.03"
        self.data_dir.mkdir()
        (self.data_dir / "F1.csv").write_text(
            "组织,科目,2027.03\nA,B,1\n", encoding="utf-8-sig")
        self.flow = {
            "navigator": {
                "source": "file", "launch_file": str(self.nav),
                "workbook_pattern": "nav*", "sheet": "nav", "version_cell": "AR45",
                "refresh_flow_macro": "Refresh_Status", "open_button_cell": "AT52",
                "open_wait_sec": 14,
            },
            "package": {
                "workbook_pattern": "pkg*", "sheet": "pkg",
                "execute_macro": "Trigger_PS", "refresh_macro": "Refresh_Sheet",
                "status_cell": "AO73", "max_wait_sec": 180,
            },
            "verify": {
                "nav_cell": "AC70", "form_row": 54, "sheet": "v",
                "header_row": 54, "indicator_rows": {"净收入": 69},
            },
            "steps": [{"form": "F1", "nav_cell": "AC57", "data": "F1.csv"}],
        }

    def tearDown(self):
        self.tmp.cleanup()

    def test_valid_flow(self):
        issues = cc.check_flow(self.flow, {"F1": KEYED}, self.data_dir)
        self.assertEqual(levels(issues), [], messages(issues))

    def test_missing_nav_file(self):
        flow = dict(self.flow)
        flow["navigator"] = dict(self.flow["navigator"], launch_file=str(self.root / "nope.xlsb"))
        issues = cc.check_flow(flow, {"F1": KEYED}, self.data_dir)
        self.assertIn("导航表不存在", messages(issues))

    def test_missing_data_dir(self):
        issues = cc.check_flow(self.flow, {"F1": KEYED}, self.root / "missing")
        self.assertIn("数据目录不存在", messages(issues))

    def test_unreferenced_form(self):
        issues = cc.check_flow(self.flow, {}, self.data_dir)
        self.assertIn("表单配置不存在", messages(issues))

    def test_config_only_skips_data(self):
        issues = cc.check_flow(self.flow, {"F1": KEYED}, self.root / "missing",
                               check_data=False)
        # 数据目录仍要求存在，但不检查文件内容
        self.assertIn("数据目录不存在", messages(issues))


if __name__ == "__main__":
    unittest.main()
