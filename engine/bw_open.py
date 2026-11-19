"""通过 SAP Analysis for Office 的 BW 平台打开 PBF 导航表。

为什么不用本地的导航表 .xlsb：
    本地文件里缓存了保存时的连接与版本上下文，用「双击打开」的方式进入时，
    AO 只对保存时那个版本稳定；切到其他版本常常出现刷新失败/版本不生效。
    从 BW 平台现取导航表，每次都是干净的服务器副本，切版本更稳定。

现场路径（本模块用 uiautomation 逐步复现）：
    ① 用 Excel 打开一个空表
    ② 文件 → Analysis → 打开工作簿 → 从 SAP Business Warehouse 平台打开工作簿
    ③ 双击 TCL BPC PRD → 登录（用户/密码）——会话仍在时这一步可能被跳过
    ④ 搜索并双击 PBF Pro - 全面预算系统流程导航
    ⑤ 之后由 orchestrator 负责版本检查/切换与「刷新状态」

实现注意：
- uiautomation 的 Control(SubName=...) 底层用 UIA Find，会漏掉 Excel 后台视图
  （NetUI）里的部分元素；这里统一改为 GetChildren 逐层遍历 + 谓词过滤。
- 菜单项要点具体的 MenuItemControl「打开工作簿」，点同名的 Text/Group 不会展开。
- 登录框、连接选择框、文档浏览框的出现顺序不固定（有缓存会话时直接进文档浏览），
  因此用状态机轮询，而不是按固定顺序等待。
"""
from __future__ import annotations

import logging
import subprocess
import tempfile
import time
import uuid
from pathlib import Path
from typing import Callable, Optional

import uiautomation as auto

from .excel import ExcelEngine, ExcelError, reset_active_excel

log = logging.getLogger(__name__)

BW_MENU_ITEM = "SAP Business Warehouse 平台打开工作簿"
OPEN_TILE = "打开工作簿"
CONNECTION = "TCL BPC PRD"
NAV_TITLE = "PBF Pro - 全面预算系统流程导航"
LOGIN_TITLE = "登录到系统"


# ---------- 基础 UI 操作 ----------

def _foreground(hwnd: int) -> None:
    try:
        import ctypes

        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, 3)                  # SW_MAXIMIZE
        user32.keybd_event(0x12, 0, 0, 0)           # 解除前台锁
        user32.keybd_event(0x12, 0, 2, 0)
        user32.SetForegroundWindow(hwnd)
    except Exception as exc:
        log.debug(f"Excel 置前失败（忽略）: {exc}")
    time.sleep(0.8)


def _click(ctrl) -> None:
    try:
        ctrl.Click(simulateMove=False)
        return
    except Exception:
        pass
    r = ctrl.BoundingRectangle
    x, y = (r.left + r.right) // 2, (r.top + r.bottom) // 2
    auto.SetCursorPos(x, y)
    time.sleep(0.15)
    auto.Click()


def _double_click(ctrl) -> None:
    try:
        ctrl.DoubleClick(simulateMove=False)
        return
    except Exception:
        pass
    _click(ctrl)
    time.sleep(0.15)
    _click(ctrl)


def _name(ctrl) -> str:
    try:
        return ctrl.Name or ""
    except Exception:
        return ""


def _aid(ctrl) -> str:
    try:
        return ctrl.AutomationId or ""
    except Exception:
        return ""


def _walk(root, pred: Callable, maxdepth: int = 14, depth: int = 0, out=None):
    """深度优先遍历控件树，返回第一个满足 pred 的控件。"""
    if out is None:
        out = []
    if depth > maxdepth or out:
        return out
    try:
        children = root.GetChildren()
    except Exception:
        return out
    for ch in children:
        try:
            if pred(ch):
                out.append(ch)
                return out
        except Exception:
            pass
        _walk(ch, pred, maxdepth, depth + 1, out)
        if out:
            return out
    return out


def _find(root, pred: Callable, timeout: float = 20.0, maxdepth: int = 14):
    deadline = time.time() + timeout
    while time.time() < deadline:
        out = _walk(root, pred, maxdepth)
        if out:
            return out[0]
        time.sleep(0.4)
    return None


def _top_windows(pid: Optional[int] = None) -> list:
    try:
        tops = auto.GetRootControl().GetChildren()
    except Exception:
        return []
    if pid is None:
        return tops
    out = []
    for w in tops:
        try:
            if w.ProcessId == pid:
                out.append(w)
        except Exception:
            continue
    return out


def _wait_top_window(pid: Optional[int], parts: tuple[str, ...],
                     timeout: float = 60.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for w in _top_windows(pid):
            name = _name(w)
            if all(p in name for p in parts):
                return w
        time.sleep(0.4)
    return None


def _browse_dialog(win, pid: Optional[int]):
    """AO 的『打开文档』窗口是 Excel 主窗口的子窗口（AutomationId=CnOpenBrowseDialog），
    不是顶层窗口；部分版本也可能浮成顶层窗口，两种都兼容。"""
    d = _find(win, lambda c: _aid(c) == "CnOpenBrowseDialog", timeout=0.1, maxdepth=12)
    if d is not None:
        return d
    return _wait_top_window(pid, ("打开文档",), timeout=0.1)


def _list_item(root, text: str):
    return _find(
        root,
        lambda c: c.ControlTypeName == "ListItemControl"
        and (text == _name(c).strip() or text in _name(c)),
        timeout=0.2, maxdepth=14)


def _close_stale_browse(win) -> None:
    """上次中断可能留下模态的『打开文档』窗口，它会阻塞整个 Excel UI，先关掉。"""
    d = _find(win, lambda c: _aid(c) == "CnOpenBrowseDialog", timeout=0.5, maxdepth=12)
    if d is None:
        return
    cancel = _find(d, lambda c: _aid(c) == "_cancelButton", timeout=1.0, maxdepth=6)
    if cancel is not None:
        _click(cancel)
    else:
        auto.SendKeys("{Esc}")
    log.warning("已关闭上次残留的『打开文档』窗口")
    time.sleep(1.0)


def _login_visible() -> bool:
    # 登录框标题是「登录到系统 TCL BPC PRD」，这里必须用“包含”匹配；
    # 用 Name= 精确匹配会永远等不到，表现为一直卡在连接/登录这一步。
    try:
        for w in auto.GetRootControl().GetChildren():
            if LOGIN_TITLE in _name(w):
                return True
    except Exception:
        pass
    try:
        return auto.WindowControl(SubName=LOGIN_TITLE, searchDepth=2).Exists(0.1)
    except Exception:
        return False


# ---------- 主流程 ----------

def _open_ui(engine: ExcelEngine, flow_cfg: dict, connection: str, nav_title: str,
             wait_sec: int) -> None:
    hwnd = int(engine._excel.Hwnd)
    try:
        import win32process

        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception:
        pid = None
    if not pid:
        raise ExcelError("无法获取 Excel 进程号，UI 自动化无法定位窗口")

    # 主线程第一次用 UIA：只初始化，不反初始化，避免影响 pywin32 已建立的 COM 环境
    auto.InitializeUIAutomationInCurrentThread()

    win = auto.ControlFromHandle(hwnd)
    if win is None:
        raise ExcelError("无法通过窗口句柄定位 Excel 主窗口")

    _foreground(hwnd)
    _close_stale_browse(win)
    auto.SendKeys("{Esc}")
    time.sleep(0.3)

    # ① 文件（后台视图）：Alt+F 最稳；失败再点「文件」选项卡
    auto.SendKeys("{Alt}f")
    time.sleep(1.5)
    analysis = _find(
        win,
        lambda c: c.ControlTypeName == "ListItemControl" and _name(c).strip() == "Analysis",
        timeout=6, maxdepth=12)
    if analysis is None:
        file_tab = _find(
            win,
            lambda c: _aid(c) == "FileTabButton" or
            (c.ControlTypeName == "ButtonControl" and "文件" in _name(c)),
            timeout=10, maxdepth=8)
        if file_tab is not None:
            log.info("Alt+F 未打开后台视图，改为点击『文件』选项卡")
            _click(file_tab)
            time.sleep(2.0)
            analysis = _find(
                win,
                lambda c: c.ControlTypeName == "ListItemControl"
                and _name(c).strip() == "Analysis",
                timeout=15, maxdepth=12)

    # ② 后台视图 → Analysis 页
    if analysis is None:
        raise ExcelError(
            "在 Excel 后台视图里找不到『Analysis』页。请确认 SAP Analysis for Office "
            "插件已加载（Excel 选项 → 加载项），并保持 Excel 窗口在前台。")
    _click(analysis)
    time.sleep(1.0)
    log.info("已进入 Analysis 页")

    # ③ 打开工作簿（必须点 MenuItem，点同名的 Group/Text 不会展开菜单）
    tile = _find(
        win,
        lambda c: c.ControlTypeName == "MenuItemControl" and _name(c).strip() == OPEN_TILE,
        timeout=20, maxdepth=12)
    if tile is None:
        raise ExcelError(f"Analysis 页找不到『{OPEN_TILE}』按钮")
    _click(tile)
    time.sleep(0.8)

    item = _find(
        win,
        lambda c: c.ControlTypeName == "MenuItemControl" and BW_MENU_ITEM in _name(c),
        timeout=20, maxdepth=14)
    if item is None:
        raise ExcelError(f"找不到菜单项『{BW_MENU_ITEM}』")
    _click(item)
    log.info("已点击『从 BW 平台打开工作簿』")

    # ④ 登录/连接/文档浏览的状态机：顺序不固定，有缓存会话时直接进文档浏览
    login_t = None
    target = None
    deadline = time.time() + max(wait_sec, 120)
    while time.time() < deadline:
        if login_t is None and _login_visible():
            log.info("检测到 AO 登录框，自动填写")
            from .bootstrap import _start_ao_login
            login_t = _start_ao_login(flow_cfg)
            if login_t is not None:
                login_t.join(timeout=min(120.0, max(30.0, deadline - time.time())))
                if login_t.is_alive():
                    raise ExcelError("AO 自动登录超时，请检查账号密码后重试")
            continue

        dlg = _browse_dialog(win, pid)
        if dlg is None:
            time.sleep(0.4)
            continue

        # 文档列表：目标导航表已经出现
        nav = _list_item(dlg, nav_title)
        if nav is not None:
            target = nav
            break

        # 连接列表：先选连接，之后可能是登录框，也可能是文档列表
        conn = _list_item(dlg, connection)
        if conn is not None:
            log.info(f"选择连接：{connection}")
            _double_click(conn)
            time.sleep(2.0)
            continue
        time.sleep(0.4)

    if target is None:
        raise ExcelError(
            f"打开 BW 工作簿超时：未在『打开文档』窗口里找到『{nav_title}』。"
            "请确认账号密码正确、网络可达，且没有被『不安全连接』等弹窗拦下。")
    _double_click(target)
    log.info(f"已选择导航表：{nav_title}")


def open_navigator_from_bw(flow_cfg: dict, wait_sec: int = 180) -> ExcelEngine:
    """用「Excel 空表 + BW 平台」打开导航表，返回已连到该 Excel 的引擎。"""
    nav = flow_cfg.get("navigator") or {}
    pattern = nav.get("workbook_pattern", "PBF Pro - 全面预算系统流程导航*")
    connection = str(nav.get("bw_connection") or CONNECTION)
    nav_title = str(nav.get("bw_nav_title") or NAV_TITLE)

    excel_exe = ExcelEngine._real_excel_path()
    if not excel_exe:
        raise ExcelError("找不到真实 Excel（EXCEL.EXE），请确认 Office 已安装")

    # 用一个空白工作簿启动 Excel（等价于双击一个 .xlsx），确保 AO 插件完整加载
    blank = Path(tempfile.gettempdir()) / f"PBF_启动空表_{uuid.uuid4().hex[:8]}.xlsx"
    from openpyxl import Workbook

    Workbook().save(blank)

    log.info("启动 Excel（空表），准备从 BW 平台打开导航表…")
    reset_active_excel()
    try:
        subprocess.Popen([excel_exe, str(blank)])
    except Exception as exc:
        blank.unlink(missing_ok=True)
        raise ExcelError(f"启动 Excel 失败：{exc}") from exc

    time.sleep(4)
    engine = ExcelEngine(timeout_sec=30)
    try:
        engine._excel.DisplayAlerts = False
    except Exception:
        pass

    try:
        _open_ui(engine, flow_cfg, connection, nav_title, wait_sec)
        engine.get_workbook(pattern, wait_sec=wait_sec)
        log.info("导航表已从 BW 平台打开，等待 AO 加载…")
        time.sleep(10)
    finally:
        # 关掉启动用的空表（只关自己建的那本，且导航表已经打开时才关）
        try:
            if engine.workbook_count() > 1:
                for wb in list(engine._excel.Workbooks):
                    if wb.Name == blank.name:
                        engine.close_workbook(wb, save=False)
        except Exception as exc:
            log.debug(f"关闭启动空表失败（忽略）: {exc}")
        try:
            blank.unlink(missing_ok=True)
        except Exception:
            pass
    return engine
