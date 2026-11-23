"""导航表步骤面板刷新检测。

背景（2026-11-23 现场）：网页清 0 运行到最后一步时，AO 会话已失效，
选中 AC70 后右侧「步骤事项」面板仍停在上一张表的清单，AT52 点下去
毫无反应，重试 3 次后报「未打开新工作簿」。修复后必须先确认面板
真的刷新、按钮对应行类型正确，会话失效则抛 SessionExpiredError
（由 run_with_recovery 强制重启 Excel 并重登）。
"""
from __future__ import annotations

import sys
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.excel import ExcelError
from engine.orchestrator import PbfFlow, SessionExpiredError


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

    def _retry(self, fn, attempts: int = 1, delay: float = 0.0):
        return fn()

    def get_workbook(self, pattern: str, wait_sec: int = 0) -> FakeWorkbook:
        return self.wb

    def open_sheet(self, wb, name: str) -> FakeSheet:
        return self.sheet

    def workbook_names(self) -> list[str]:
        return list(self.names)

    def read(self, sheet: FakeSheet, cell: str):
        return sheet.value(cell)

    def select(self, sheet: FakeSheet, cell: str) -> None:
        sheet.select(cell)

    def run_macro(self, name: str, wb=None, *args):
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
        "package": {"workbook_pattern": "PBF Pro运行包*", "sheet": "pkg"},
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
        # 先得到 AC62 正确面板的指纹（模拟上次运行留下的状态）
        flow1, _, sheet1 = make_flow()
        flow1.engine.select(sheet1, "AC62")
        fingerprint = flow1._panel_fingerprint()

        flow2, _, sheet2 = make_flow(refresh_on_select=False)
        # 新进程启动时面板已经停在 AC62 上
        sheet2.panel = {r: list(v) for r, v in sheet2.panels["AC62"].items()}
        flow2._panel_cache[flow2._panel_cache_key("AC62")] = fingerprint

        name = flow2.open_via_nav("AC62", expect="package")
        self.assertIn("运行包", name)

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


if __name__ == "__main__":
    unittest.main()
