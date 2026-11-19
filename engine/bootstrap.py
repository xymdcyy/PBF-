"""Excel 启动/恢复：确保 Excel 进程与 PBF 导航表处于可用状态。

一键执行和「运行中断后自动恢复」都走这里：
- 已就绪 → 直接返回
- 进程残留但 COM 断开 → 结束残留进程（kill_stale=True 时）
- 否则用真实 Excel 的 EXE 以文档参数启动（等价于双击），轮询到工作簿出现
"""
from __future__ import annotations

import logging
import subprocess
import threading
import time
from pathlib import Path

from .excel import ExcelEngine, ExcelError, reset_active_excel

log = logging.getLogger(__name__)


def excel_pids() -> list[int]:
    """当前 Excel 进程 PID 列表（tasklist 解析）。"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return []
    pids: list[int] = []
    for line in out.splitlines():
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE" and parts[1].isdigit():
            pids.append(int(parts[1]))
    return pids


def kill_excel() -> None:
    log.warning("结束无响应的 Excel 进程…")
    try:
        subprocess.run(["taskkill", "/IM", "EXCEL.EXE", "/F"],
                       capture_output=True, text=True, timeout=30)
    except Exception as exc:
        log.warning(f"结束 Excel 失败（忽略）: {exc}")
    reset_active_excel()  # 进程已结束，清掉缓存的 COM 引用
    time.sleep(3)


def _start_ao_login(flow_cfg: dict):
    """后台线程自动登录 AO（返回线程对象，未配置账号时返回 None）。

    AO 连接 BPC 时弹出的登录框是模态的，会阻塞 Excel 的 COM 消息泵，导致主线程的
    Workbooks 枚举卡住。这里用 uiautomation 在后台线程填登录框（不依赖 Excel COM），
    登录完成后模态框消失，主线程自然恢复。
    """
    try:
        from .orchestrator import login_config
        from .ao_login import AoLogin
    except Exception:
        return None
    login_cfg = login_config(flow_cfg)
    if not login_cfg.get("username"):
        return None

    def _login() -> None:
        # uiautomation 需要在线程内初始化 COM（CoInitialize + 建 UIAutomation 对象），
        # 否则报 "尚未调用 CoInitialize"；用库自带的初始化器包一层。
        import uiautomation as auto
        with auto.UIAutomationInitializerInThread():
            try:
                ok = AoLogin(
                    username=login_cfg["username"],
                    password=login_cfg.get("password", ""),
                    client=str(login_cfg.get("client", "333")),
                    lang=str(login_cfg.get("lang", "zh")),
                ).ensure(float(login_cfg.get("timeout", 120)), expect_login=True)
                if not ok:
                    log.warning("后台自动登录超时，请检查账号密码")
            except Exception as exc:
                log.warning(f"后台自动登录异常: {exc}")

    t = threading.Thread(target=_login, daemon=True)
    t.start()
    return t


def ensure_excel_ready(flow_cfg: dict, kill_stale: bool = False, wait_sec: int = 180):
    """确保真实 Excel 已运行且导航表已打开，否则自动打开。

    返回可复用的 ExcelEngine（所有成功路径），调用方可复用同一实例，避免重连到别的 Excel。
    打开导航表用真实 Excel 的 EXE 以文档参数启动（等价于双击文件），完整加载 AO 并连接
    BPC；若用 COM 的 Workbooks.Open，AO 常不初始化、进不了 PBF。
    """
    nav = flow_cfg.get("navigator") or {}
    pattern = nav.get("workbook_pattern", "PBF Pro - 全面预算系统流程导航*")
    path = Path(nav["launch_file"]) if nav.get("launch_file") else None

    # 1) 已经就绪（真实 Excel 里导航表已打开）——仅当确有 Excel 进程时才去连，
    #    否则 ExcelEngine 的 Dispatch 会拉起一个空 Excel，还得再关掉，反而制造竞态。
    if excel_pids():
        try:
            engine = ExcelEngine(timeout_sec=3)
            engine.find_workbook(pattern)
            return engine
        except ExcelError:
            pass

    # 2) 清理无窗口的残留实例（有 Excel 进程但没打开导航表；COM 可能连到它，
    #    后续宏会全部失败）。没有 Excel 进程时直接跳过，避免 Dispatch 又启动一个。
    if excel_pids():
        try:
            engine = ExcelEngine(timeout_sec=3)
            if engine.workbook_count() == 0:
                engine._excel.Quit()
                reset_active_excel()
                log.warning("已关闭无窗口的残留 Excel 实例")
                time.sleep(3)
        except ExcelError:
            pass

    # 3) 断开重连时：残留进程可能已经僵死，先结束掉
    if kill_stale and excel_pids():
        kill_excel()

    # 4) 打开导航表
    #    source=bw（默认）：启动空白 Excel → 文件 → Analysis → 从 BW 平台现取导航表。
    #      本地 .xlsb 会带缓存的连接/版本上下文，切版本容易出错；BW 每次都是新副本。
    #    source=file：直接打开 flow.yaml 里配置的本地文件（旧方式，保留兜底）。
    source = str(nav.get("source") or "bw").lower()
    if source == "bw":
        from .bw_open import open_navigator_from_bw

        return open_navigator_from_bw(flow_cfg, wait_sec=wait_sec)

    if not path or not path.exists():
        raise ExcelError(
            f"导航表未打开，且路径无效：{path}；请检查 config/flow.yaml 的 navigator.launch_file")
    excel_exe = ExcelEngine._real_excel_path()
    if not excel_exe:
        raise ExcelError("找不到真实 Excel（EXCEL.EXE），无法打开导航表，请确认 Office 已安装")
    log.info(f"正在打开导航表：{path}")
    reset_active_excel()
    try:
        subprocess.Popen([excel_exe, str(path)])
    except Exception as exc:
        raise ExcelError(f"启动 Excel 打开导航表失败：{exc}") from exc
    # 后台登录：模态登录框会阻塞 Excel COM，用 uiautomation 线程先把它填掉
    login_t = _start_ao_login(flow_cfg)
    # 等 Excel 进程起来（避免下面的 Dispatch 又拉起第二个实例），再连 COM
    time.sleep(3)
    engine = ExcelEngine(timeout_sec=20)
    try:
        engine._excel.DisplayAlerts = False
    except Exception:
        pass

    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(3)
        try:
            engine.find_workbook(pattern)
            log.info("导航表已打开，等待 AO 加载…")
            time.sleep(10)
            if login_t is not None:
                login_t.join(timeout=60)  # 等登录结束，保证后续切版本/刷新时已连上 BPC
            return engine
        except ExcelError:
            continue
    raise ExcelError(f"等待导航表打开超时（{wait_sec}s）")
