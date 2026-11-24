"""导航表步骤面板刷新检测。

背景（2026-11-23 现场）：网页清 0 运行到最后一步时，AO 会话已失效，
选中 AC70 后右侧「步骤事项」面板仍停在上一张表的清单，AT52 点下去
毫无反应，重试 3 次后报「未打开新工作簿」。修复后必须先确认面板
真的刷新、按钮对应行类型正确，会话失效则抛 SessionExpiredError
（由 run_with_recovery 强制重启 Excel 并重登）。
"""
from __future__ import annotations

import fnmatch
import sys
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError
from engine.orchestrator import PbfFlow, SessionExpiredError, is_session_expired


class FakeRange:
    def __init__(self, sheet: "FakeSheet", addr: str) -> None:
        self.sheet = sheet
        self.addr = addr

    def Select(self) -> None:      # noqa: N802 - 模仿 Excel COM 命名
        self.sheet.select(self.addr)

    @property
    def Value2(self):              # noqa: N802
        return self.sheet.value(self.addr)


class FakeWorkbook:
    def __init__(self, name: str) -> None:
        self.Name = name

    def Activate(self) -> None:     # noqa: N802
        return None


class FakePkgSheet:
    """运行包执行清单：AB73 步骤名 / AB60 版本 / AO73 状态 / AW73 时长。"""

    def __init__(self, cells: dict) -> None:
        self.cells = dict(cells)

    def value(self, addr: str):
        return self.cells.get(addr)


class FakeSheet:
    """导航表：面板按选中业务单元静态写入，模拟真实 VBA 行为。"""

    def __init__(self, engine: "FakeEngine", panels: dict[str, dict[int, tuple]],
                 refresh_on_select: bool = True) -> None:
        self.engine = engine
        self.panels = panels
        self.refresh_on_select = refresh_on_select
        self.panel: dict[int, list] = {}
        self.version = "滚动预测2026年11月财总版"
        self.selection_log: list[str] = []
        self.btn = "AT52"
        # 真实 AO 打开的工作簿名带时间戳（…(17-55-41).xlsb），每次都不一样；
        # 若这里重复追加同名，"新增工作簿"按集合差集判断会永远为空。
        self.open_count = 0
        self.clicks_blocked = False

    # COM 形状
    def Activate(self) -> None:     # noqa: N802
        return None

    def Range(self, addr: str) -> FakeRange:   # noqa: N802
        return FakeRange(self, addr)

    # 行为
    def select(self, addr: str) -> None:
        self.selection_log.append(addr)
        if addr == "A1":
            return
        if addr in self.panels:
            if self.refresh_on_select:
                self.panel = {r: list(v) for r, v in self.panels[addr].items()}
            return
        if (addr == self.btn and not self.clicks_blocked
                and self.panel.get(52, [None] * 5)[4] == "打开"):
            self.open_count += 1
            self.engine.names.append(
                f"PBF Pro运行包执行及日志查看清单 (9-{self.open_count:02d}).xlsb")

    def value(self, addr: str):
        if addr == "AP48:AU60":
            return tuple(
                tuple(self.panel.get(r, [None] * 6)[:6]) for r in range(48, 61)
            )
        if addr == "AR45":
            return self.version
        if addr.startswith("AT"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[4] if row else None
        if addr.startswith("AR"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[2] if row else None
        if addr.startswith("AP"):
            r = int(addr[2:])
            row = self.panel.get(r)
            return row[0] if row else None
        return None


class FakeEngine:
    def __init__(self, sheet: FakeSheet, wb: FakeWorkbook) -> None:
        self.sheet = sheet
        self.wb = wb
        self.names = [wb.Name]
        self.pid = 4321
        self.macro_error: Exception | None = None
        self.pkg_sheet: FakePkgSheet | None = None
        self.closed: list[str] = []
        self.macros_run: list[str] = []

    def _retry(self, fn, attempts: int = 1, delay: float = 0.0):
        return fn()

    def _find_wb(self, pattern: str) -> FakeWorkbook:
        for name in self.names:
            if fnmatch.fnmatchcase(name, pattern):
                if "运行包" in name:
                    # 运行包清单每次打开都是新副本（名字带时间戳）
                    return FakeWorkbook(name)
                return self.wb
        raise ExcelError(f"没找到匹配 {pattern} 的工作簿")

    def get_workbook(self, pattern: str, wait_sec: int = 0) -> FakeWorkbook:
        return self._find_wb(pattern)

    def find_workbook(self, pattern: str) -> FakeWorkbook:
        return self._find_wb(pattern)

    def open_sheet(self, wb, name: str):
        if "运行包" in getattr(wb, "Name", "") and self.pkg_sheet is not None:
            return self.pkg_sheet
        return self.sheet

    def workbook_names(self) -> list[str]:
        return list(self.names)

    def close_workbook(self, wb, save: bool = False) -> None:
        self.closed.append(wb.Name)
        self.names = [n for n in self.names if n != wb.Name]

    def read(self, sheet, cell: str):
        return sheet.value(cell)

    def select(self, sheet: FakeSheet, cell: str) -> None:
        sheet.select(cell)

    def run_macro(self, name: str, wb=None, *args):
        self.macros_run.append(name)
        if self.macro_error is not None:
            raise self.macro_error
        return None


def make_flow(refresh_on_select: bool = True):
    panels = {
        "AC62": {
            52: ["PS_经营模型_收入_预测_滚动预测收入计算", "ZPS_1", "运行包", "收入计算", "打开"],
            53: ["PS_经营模型_收入_预测_计算到销售组织的成本", "ZPS_2", "运行包", "成本计算", "打开"],
        },
        "AC70": {
            52: ["YC_CBG损益查询表", "ZWB_1", "表单", "损益查询", "打开"],
            53: ["PS_计算损益_滚动预测", "ZPS_9", "运行包", "计算滚动预测损益", "打开"],
        },
        "AC55": {
            52: ["YCSRW026_价格&销量&加权CT编制表", "ZWB_X", "表单", "价格销量", "打开"],
        },
    }
    wb = FakeWorkbook("PBF Pro - 全面预算系统流程导航.xlsb")
    sheet = FakeSheet(None, panels, refresh_on_select=refresh_on_select)
    engine = FakeEngine(sheet, wb)
    sheet.engine = engine
    # 初始面板：上一张表的清单，AT52 仍是"打开"——最迷惑人的状态
    sheet.panel = {52: ["上一张表", "ZWB_PREV", "表单", "上一张表", "打开"]}

    cfg = {
        "navigator": {
            "workbook_pattern": "PBF Pro*导航*",
            "sheet": "nav",
            "open_button_cell": "AT52",
            "refresh_flow_macro": "Refresh_Status",
            "open_wait_sec": 0.2,
            "panel_range": "AP48:AU60",
            "panel_wait_sec": 0.2,
            "version_cell": "AR45",
        },
        "package": {
            "workbook_pattern": "PBF Pro运行包执行及日志查看清单*",
            "sheet": "运行包执行及日志查看清单",
            "execute_macro": "Trigger_PS",
            "refresh_macro": "Refresh_Sheet",
            "version_cell": "AB60",
            "status_cell": "AO73",
            "step_cell": "AB73",
            "duration_cell": "AW73",
            "poll_interval_sec": 0,
            "max_wait_sec": 1,
        },
        "verify_writeback": {},
    }
    flow = PbfFlow(engine, cfg, {}, Path("data_2027.02_clear"))
    return flow, engine, sheet


class PanelRefreshTests(unittest.TestCase):
    def test_stale_panel_raises_session_expired(self) -> None:
        flow, engine, sheet = make_flow(refresh_on_select=False)

        def dead(attempts=0, delay=0.0):
            raise SessionExpiredError("刷新流程失败：AO/BPC 会话可能已过期")

        flow.refresh_flow_retry = dead
        with self.assertRaises(SessionExpiredError):
            flow.open_via_nav("AC70", open_cell="AT52", expect="package")
        # 必须先点过 A1 强制 SelectionChange（重复选中同一格不会触发事件）
        self.assertEqual(sheet.selection_log[:2], ["A1", "AC70"])

    def test_panel_refresh_opens_package(self) -> None:
        flow, engine, sheet = make_flow()
        name = flow.open_via_nav("AC62", open_cell="AT52", expect="package")
        self.assertIn("运行包", name)
        self.assertIn(flow._panel_cache_key("AC62"), flow._panel_cache)

    def test_wrong_row_type_is_rejected(self) -> None:
        # AC70 的面板里 AT52 是表单；运行包实际在 AT53。行号配错必须报错，
        # 不能"打开"了表单还以为在跑运行包。
        flow, _, _ = make_flow()
        with self.assertRaises(ExcelError) as ctx:
            flow.open_via_nav("AC70", open_cell="AT52", expect="package")
        self.assertIn("行类型", str(ctx.exception))
        self.assertIn("表单", str(ctx.exception))

    def test_form_open_still_works(self) -> None:
        flow, engine, sheet = make_flow()
        sheet.panels["AC55"][52][4] = "打开"
        name = flow.open_via_nav("AC55", expect="form")
        self.assertIn("运行包", name)   # fake 打开后统一返回运行包清单名

    def test_refresh_flow_maps_0x800A9C68(self) -> None:
        flow, engine, _ = make_flow()
        engine.macro_error = ExcelError(
            "macro 'x'!Refresh_Status: (-2147352567, '发生意外。', (0, None, None, None, 0, -2146788248), None)")
        with self.assertRaises(SessionExpiredError):
            flow.refresh_flow()


class PanelCacheTests(unittest.TestCase):
    def test_reopening_same_nav_cell_not_judged_stale(self) -> None:
        """同一单元格重开（落库校验）时面板内容与上次一致，应命中缓存。"""
        flow, engine, sheet = make_flow()
        flow.open_via_nav("AC55", expect="form")
        # 再选同一行：A1 不改变面板 → 选中 AC55 后面板内容与 before 相同
        name = flow.open_via_nav("AC55", expect="form")
        self.assertIn("运行包", name)

    def test_panel_already_on_target_hits_persisted_cache(self) -> None:
        """上次运行把面板停在目标行（如手工探测后），新进程不应误判会话失效。

        现场 2026-11-23：清 0 续跑时面板恰好停在 AC70，选中前后内容一致，
        被误判为"面板未刷新"而强制重启 Excel——本次修复要杜绝这种误判。
        """
        flow1, _, sheet1 = make_flow()
        flow1.engine.select(sheet1, "AC62")
        fingerprint = flow1._panel_fingerprint()

        flow2, _, sheet2 = make_flow(refresh_on_select=False)
        # 新进程启动时面板已经停在 AC62 上
        sheet2.panel = {r: list(v) for r, v in sheet2.panels["AC62"].items()}
        flow2._panel_cache[flow2._panel_cache_key("AC62")] = fingerprint

        name = flow2.open_via_nav("AC62", expect="package")
        self.assertIn("运行包", name)

    def test_panel_already_on_target_without_cache_accepted_by_row_type(self) -> None:
        """冷启动（无缓存）+ 面板停在目标行：刷新成功过一轮后应按行类型放行。

        原缺陷：新进程没有持久化缓存，面板恰好停在目标行、内容不再变化，
        第 2 轮重试会把"没变化"判成会话失效，白重启一次 Excel 并重新登录。
        现在第 1 轮仍交给刷新流程分辨，刷新成功过（会话活着）之后，
        按钮行行类型与期望一致就接受静止面板。
        """
        flow, _, sheet = make_flow(refresh_on_select=False)
        # 面板本来就停在 AC62 的正确内容上，且没有持久化缓存
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC62"].items()}

        name = flow.open_via_nav("AC62", open_cell="AT52", expect="package")
        self.assertIn("运行包", name)
        # 静止面板被采用后也应落库，下次可直接命中缓存
        self.assertIn(flow._panel_cache_key("AC62"), flow._panel_cache)

    def test_stale_static_panel_with_wrong_row_type_still_rejected(self) -> None:
        """静止面板若停在别的行上，行类型对不上仍要拦下，不能被静态放行绕过。"""
        flow, _, sheet = make_flow(refresh_on_select=False)
        # 面板停在 AC70 的内容：AT52 行类型是「表单」，与期望的「运行包」不符
        sheet.panel = {r: list(v) for r, v in sheet.panels["AC70"].items()}
        with self.assertRaises(SessionExpiredError):
            flow.open_via_nav("AC62", open_cell="AT52", expect="package")

    def test_cached_panel_click_no_response_raises_session_expired(self) -> None:
        """面板是"已知正确内容"但点击毫无反应：会话大概率已死，必须走恢复。

        与原缺陷不同的是：这里不是面板停在别人身上（那会用行类型校验拦下），
        而是面板内容正确、按钮也是"打开"，点下去却什么都没发生。
        """
        flow1, _, sheet1 = make_flow()
        flow1.engine.select(sheet1, "AC62")
        fingerprint = flow1._panel_fingerprint()

        flow2, _, sheet2 = make_flow(refresh_on_select=False)
        sheet2.panel = {r: list(v) for r, v in sheet2.panels["AC62"].items()}
        sheet2.clicks_blocked = True
        flow2._panel_cache[flow2._panel_cache_key("AC62")] = fingerprint
        flow2.nav_cfg["open_wait_sec"] = 0.01

        with self.assertRaises(SessionExpiredError):
            flow2.open_via_nav("AC62", expect="package")

    def test_cache_persisted_across_processes(self) -> None:
        import json
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            cache_file = Path(td) / "nav_panel_cache.json"
            flow, _, _ = make_flow()
            flow.persist_panel_cache = True
            flow._panel_cache_file = cache_file
            flow.open_via_nav("AC62", expect="package")

            key = flow._panel_cache_key("AC62")
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            self.assertIn(key, data)

            flow2, _, _ = make_flow()
            flow2._panel_cache_file = cache_file
            flow2._load_panel_cache()
            self.assertEqual(flow2._panel_cache.get(key), data[key])


class RunPackageTests(unittest.TestCase):
    """运行包 AB73 步骤名校验：面板停在别的同类型行上时绝不执行错包。

    评审 2026-11-24：accept_static 区分不了「面板停在目标行」和「面板停在
    另一个同类型行」，点下去会开到错误的运行包上下文。清单工作簿名不随包变，
    只能靠 AB73 步骤名与 flow.yaml 配置的期望名核对。
    """

    PKG_CELLS = {
        "AB60": "滚动预测2026年11月财总版",
        "AB73": "收入计算",
        "AO73": "已完成",
        "AW73": "0:01:23",
    }

    def test_step_name_match_runs_package(self) -> None:
        flow, engine, _ = make_flow()
        engine.pkg_sheet = FakePkgSheet(dict(self.PKG_CELLS))
        r = flow.run_package("AC62", 52, expected_name="收入计算")
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["package"], "收入计算")
        self.assertIn("Trigger_PS", engine.macros_run)
        self.assertIn("Refresh_Sheet", engine.macros_run)
        # 跑完关闭清单
        self.assertEqual(engine.closed, ["PBF Pro运行包执行及日志查看清单 (9-01).xlsb"])

    def test_step_name_mismatch_raises_without_executing(self) -> None:
        flow, engine, _ = make_flow()
        engine.pkg_sheet = FakePkgSheet(dict(self.PKG_CELLS))   # 清单显示「收入计算」
        with self.assertRaises(ExcelError) as ctx:
            flow.run_package("AC62", 52, expected_name="成本计算")
        self.assertIn("步骤名不符", str(ctx.exception))
        self.assertIn("成本计算", str(ctx.exception))
        # 核对不过就中止：Execute 宏一个都不能跑
        self.assertNotIn("Trigger_PS", engine.macros_run)
        self.assertEqual(engine.closed, ["PBF Pro运行包执行及日志查看清单 (9-01).xlsb"])

    def test_legacy_list_packages_skip_verification(self) -> None:
        # 旧行号列表写法（无期望包名）：不校验，行为同前
        flow, engine, _ = make_flow()
        engine.pkg_sheet = FakePkgSheet(dict(self.PKG_CELLS))
        r = flow.run_package("AC62", 52)
        self.assertTrue(r["ok"], r)


class SessionExpiredDetectionTests(unittest.TestCase):
    """is_session_expired：按文本特征码识别宏里直接抛出的会话失效。

    run_with_recovery 只对 SessionExpiredError 和「Excel 断开」做恢复；
    trigger_calc/run_package 的宏不走 refresh_flow 的翻译，0x800A9C68
    会裸抛上来——恢复前必须先把这类错误认出来。
    """

    def test_macro_session_codes_are_detected(self) -> None:
        self.assertTrue(is_session_expired(ExcelError(
            "macro 'x'!Execute: (-2147352567, '发生意外。', "
            "(0, None, None, None, 0, -2146788248), None)")))
        self.assertTrue(is_session_expired(Exception("0x800A9C68")))
        self.assertTrue(is_session_expired(
            SessionExpiredError("面板未刷新 (0x800A9C68)")))

    def test_generic_wrapper_codes_are_not_detected(self) -> None:
        # 「发生意外」/-2147352567 是通用 DISP_E_EXCEPTION 包装：
        # 版本与期间不匹配等普通错误也长这样，不能据此强制重启 Excel
        self.assertFalse(is_session_expired(ExcelError("发生意外。")))
        self.assertFalse(is_session_expired(ExcelError(
            "(-2147352567, '发生意外。', (0, None, None, None, 0, -2147352567), None)")))
        self.assertFalse(is_session_expired(ExcelError("未打开新工作簿")))
        self.assertFalse(is_session_expired(None))

    def test_session_expired_error_instance_without_code(self) -> None:
        # 纯文本的 SessionExpiredError 靠 run_with_recovery 的 isinstance
        # 分支识别，is_session_expired 只管文本特征——两者是互补关系
        self.assertFalse(is_session_expired(SessionExpiredError("面板一直没有刷新")))


if __name__ == "__main__":
    unittest.main()
