"""Excel 启动/恢复：确保 Excel 进程与 PBF 导航表处于可用状态。

一键执行和「运行中断后自动恢复」都走这里：
- 已就绪 → 直接返回
- 进程残留但 COM 断开 → 结束残留进程（kill_stale=True 时）
- 否则用真实 Excel 的 EXE 以文档参数启动（等价于双击），轮询到工作簿出现
"""
from __future__ import annotations

import json
import logging
import os
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path

from .cancel import raise_if_cancelled
from .excel import ExcelEngine, ExcelError, reset_active_excel
from paths import app_root

log = logging.getLogger(__name__)

# 复用已打开的 Excel 时守候登录框的默认时长（可用 flow.yaml 的
# login.reuse_wait_sec 覆盖）。AO 加载约 10s 后才弹登录框；已用缓存凭据
# 登录时框不会出现，等满 60s 会让每次复用都空等一分钟，故默认只等 15s。
REUSE_LOGIN_WAIT_SEC = 15

# 本工具新启动 Excel 时，登录框始终没出现的宽限时长（login.no_box_ok_sec）。
# 超过后按"已用缓存凭据登录"继续，不再空等满 60s 后误报"AO 自动登录失败"。
# 真正的会话失效会由后续宏操作暴露，届时提示手动登录即可。
LAUNCH_NO_BOX_OK_SEC = 30

# 登录成功 ≠ AO 就绪：加载项在登录框关闭后会自动刷新所有数据源（现场实测会
# 占住 Excel UI 线程 30s 级）。这期间调 SAPSetVariable / PauseVariableSubmit
# 会把加载项打崩，弹出「SAP BI 加载项已断开所有数据源连接」+ 空引用异常，
# 之后所有数据源都不可用。先等 refresh 启动（settle_start_sec），再用普通
# Excel 属性调用当空闲屏障——UI 线程忙时该调用会阻塞到刷新结束。
LOGIN_SETTLE_START_SEC = 5
LOGIN_IDLE_TIMEOUT_SEC = 120

# Quit 后等待进程真正退出的时长；超时视为 Quit 被外部 COM 引用忽略，强制结束。
QUIT_VERIFY_SEC = 8
KILL_VERIFY_SEC = 8

# Excel 归属记录：只有记录为 pbf 的实例才允许自动 Quit/结束。
OWNER_SOURCE_PBF = "pbf"
OWNER_SOURCE_REUSED = "reused"


def owner_file() -> Path:
    """归属记录文件；PBF_EXCEL_OWNER_FILE 可覆盖（测试用）。"""
    override = os.environ.get("PBF_EXCEL_OWNER_FILE")
    if override:
        return Path(override)
    return app_root() / "logs" / "excel_owner.json"


def _process_created(pid: int) -> int | None:
    """进程创建时间（epoch 秒）。用于识别 PID 复用，避免误关别的 Excel。"""
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        try:
            created = win32process.GetProcessTimes(handle)["CreationTime"]
        finally:
            win32api.CloseHandle(handle)
        return int(created.timestamp())
    except Exception:
        return None


def _raw_owner(path: Path | None = None) -> dict | None:
    path = path or owner_file()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def write_excel_owner(pid: int | None, navigator: str,
                      source: str = OWNER_SOURCE_PBF,
                      path: Path | None = None) -> None:
    """记录"这个 PID 的 Excel 是本工具启动的"。

    写失败不影响主流程——归属记录只是清理时的依据，丢了就退化成"不自动
    关闭"，绝不会误关用户的 Excel。
    """
    if pid is None:
        return
    path = path or owner_file()
    record = {
        "pid": int(pid),
        "created": _process_created(int(pid)),
        "navigator": str(navigator),
        "source": source,
        "recorded_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(record, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    except OSError as exc:
        log.warning(f"写入 Excel 归属记录失败（不影响运行）: {exc}")


def read_excel_owner(path: Path | None = None) -> dict | None:
    """读取归属记录并校验：进程已退出 / PID 已被复用 → 丢弃旧记录返回 None。"""
    path = path or owner_file()
    record = _raw_owner(path)
    if not record:
        return None
    try:
        pid = int(record["pid"])
    except (KeyError, TypeError, ValueError):
        clear_excel_owner(path=path)
        return None
    if pid not in excel_pids():
        clear_excel_owner(path=path)
        return None
    created = record.get("created")
    if created is not None:
        current = _process_created(pid)
        if current is not None and int(created) != current:
            log.warning(f"Excel 归属记录 PID={pid} 已被其它进程复用，丢弃旧记录")
            clear_excel_owner(path=path)
            return None
    return record


def clear_excel_owner(pid: int | None = None, path: Path | None = None) -> None:
    """删除归属记录；给定 pid 时只在记录确实指向该 PID 时删除。"""
    path = path or owner_file()
    if pid is not None:
        record = _raw_owner(path)
        try:
            if record is None or int(record.get("pid")) != int(pid):
                return
        except (TypeError, ValueError):
            return
    try:
        path.unlink()
    except FileNotFoundError:
        pass
    except OSError as exc:
        log.debug(f"删除 Excel 归属记录失败: {exc}")


def _login_seconds(flow_cfg: dict, key: str, default: float, label: str) -> float:
    """读取 flow.yaml login.<key> 的正数秒数；非法值告警并回退默认。

    守候时长只影响"等多久才认定没有登录框"，一个笔误不该悄悄改变行为，
    更不该让流程起不来。
    """
    login_cfg = flow_cfg.get("login")
    raw = login_cfg.get(key) if isinstance(login_cfg, dict) else None
    if raw is None:
        return default
    if isinstance(raw, bool):
        # bool 是 int 子类，True 会 float 成 1.0——YAML 误写 true 就成了"只等 1 秒"
        log.warning("%s=%r 不是数字，回退默认 %ss", label, raw, default)
        return default
    try:
        val = float(raw)
    except (TypeError, ValueError):
        log.warning("%s=%r 不是数字，回退默认 %ss", label, raw, default)
        return default
    if val <= 0:
        log.warning("%s=%s ≤ 0 无意义，回退默认 %ss", label, raw, default)
        return default
    return val


def _reuse_login_wait_sec(flow_cfg: dict) -> float:
    """复用 Excel 时守候登录框的秒数：flow.yaml login.reuse_wait_sec，默认 15。"""
    return _login_seconds(flow_cfg, "reuse_wait_sec", REUSE_LOGIN_WAIT_SEC,
                          "login.reuse_wait_sec")


def launch_no_box_ok_sec(flow_cfg: dict) -> float:
    """新启动 Excel 后多久没出现登录框就按"已登录"继续（login.no_box_ok_sec）。"""
    return _login_seconds(flow_cfg, "no_box_ok_sec", LAUNCH_NO_BOX_OK_SEC,
                          "login.no_box_ok_sec")


def _sleep_cancellable(seconds: float) -> None:
    """可被取消标志打断的等待（启动阶段取消要能及时退出）。"""
    end = time.time() + max(seconds, 0.0)
    while True:
        raise_if_cancelled()
        left = end - time.time()
        if left <= 0:
            return
        time.sleep(min(1.0, left))


def wait_ao_login_idle(engine: ExcelEngine, flow_cfg: dict) -> None:
    """等登录后 AO 自动刷新收尾，再让调用方去切版本/跑宏。

    只做普通 Excel 调用（Workbooks.Count），不碰 AO 变量——加载项正在刷新时
    任何 SAP* 调用都可能把它打崩；这个调用会被 Excel UI 线程排队，刷新结束才
    返回，正好当空闲屏障。登录框刚关闭时刷新可能还没启动，所以先静置一小段。
    """
    start = _login_seconds(flow_cfg, "settle_start_sec", LOGIN_SETTLE_START_SEC,
                           "login.settle_start_sec")
    timeout = _login_seconds(flow_cfg, "idle_timeout_sec", LOGIN_IDLE_TIMEOUT_SEC,
                             "login.idle_timeout_sec")
    log.info("AO 登录成功，等待其自动刷新收尾（先等 %.0fs，空闲屏障最多 %.0fs）…",
             start, timeout)
    _sleep_cancellable(start)
    if engine.wait_ui_idle(timeout_sec=timeout):
        log.info("AO 自动刷新已收尾，继续后续操作")
    else:
        log.warning("等待 AO 刷新收尾超时（%.0fs），仍继续；若随后报加载项断连，"
                    "请在 Excel 里点「重新启动」或手动切换版本", timeout)
    engine.login_settled = True


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


def kill_excel(pid: int | None = None) -> None:
    """结束 Excel 进程。

    pid 给定 → 只结束该进程（断线恢复的正确做法）；
    pid 为 None → 结束所有 EXCEL.EXE（仅保留给显式的人工命令，自动流程不要用）。
    """
    if pid is None:
        log.warning("结束所有 Excel 进程…")
        cmd = ["taskkill", "/IM", "EXCEL.EXE", "/F"]
    else:
        log.warning(f"结束 Excel 进程 PID={pid}…")
        cmd = ["taskkill", "/PID", str(pid), "/T", "/F"]
    try:
        subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except Exception as exc:
        log.warning(f"结束 Excel 失败（忽略）: {exc}")
    reset_active_excel()  # 进程已结束，清掉缓存的 COM 引用
    if pid is not None:
        clear_excel_owner(pid)
    time.sleep(3)


def _wait_pid_exit(pid: int, timeout: float) -> bool:
    """等某个 Excel PID 从进程表消失；超时返回 False。"""
    deadline = time.time() + timeout
    while True:
        if pid not in excel_pids():
            return True
        if time.time() >= deadline:
            return False
        time.sleep(0.5)


def quit_excel_verified(engine: ExcelEngine, why: str = "") -> bool:
    """Quit 并确认进程真的退出；Quit 被忽略时强制结束这个 PID。

    Excel 在有外部 COM 引用时可能不理会 Quit（进程继续以无窗口状态存活）。
    之前只调用 Quit 就记"已关闭"，随后新开的导航表经单实例机制落回这个
    僵尸实例，AO 登录框弹不出来——表现为"AO 又没了"。这里必须验证。

    只操作 engine 实际连接的 PID，绝不碰其它 Excel。
    """
    pid = getattr(engine, "pid", None)
    label = f"（{why}）" if why else ""
    try:
        engine._excel.Quit()
    except Exception as exc:
        log.warning(f"Quit 调用异常{label}: {exc}")
    reset_active_excel()
    if pid is None:
        log.warning(f"Quit 后无法确认 Excel 是否退出{label}（PID 未知）")
        return False
    if _wait_pid_exit(pid, QUIT_VERIFY_SEC):
        clear_excel_owner(pid)
        log.info(f"已关闭 Excel PID={pid}{label}")
        return True
    log.warning(f"Quit 未生效（Excel PID={pid} 仍在运行{label}），强制结束该进程…")
    kill_excel(pid)
    if _wait_pid_exit(pid, KILL_VERIFY_SEC):
        log.info(f"已强制结束 Excel PID={pid}{label}")
        return True
    log.error(f"强制结束 Excel PID={pid} 失败，进程仍在运行{label}")
    return False


def _ensure_visible(engine: ExcelEngine) -> None:
    """确保实例可见；不可见时 AO 登录框可能根本弹不出来。"""
    fn = getattr(engine, "ensure_visible", None)
    if fn is not None:
        fn()


def _record_owner(engine: ExcelEngine, navigator: str,
                  launched_pids: set[int]) -> str:
    """登记实例归属：本函数启动的新 PID → pbf，复用已有实例 → reused。

    已有的 pbf 记录不被降级：本工具上一轮启动的 Excel 这一轮被复用，仍是
    我们的实例，下一轮可以安全清理。
    """
    pid = getattr(engine, "pid", None)
    if pid is None:
        return ""
    existing = read_excel_owner()
    if existing and existing.get("pid") == pid \
            and existing.get("source") == OWNER_SOURCE_PBF:
        return OWNER_SOURCE_PBF
    source = OWNER_SOURCE_PBF if pid not in launched_pids else OWNER_SOURCE_REUSED
    write_excel_owner(pid, navigator, source)
    if source == OWNER_SOURCE_REUSED:
        log.info(f"Excel PID={pid} 是复用已有实例（单实例机制），不视为本工具所有")
    return source


class LoginHandle:
    """后台登录线程的结果句柄。

    outcome: None=线程仍在进行 | True=登录成功 | False=登录失败。
             判定登录是否完成只看 outcome（"登录成功信号"），不能拿
             "线程已启动/已结束"当成功——密码错误时线程照样退出，若被
             当成已登录，后续 COM 调用会被模态登录框永久阻塞。
    saw_box: 线程期间是否见过登录框。False+失败=从未弹框（可能已登录或
             AO 加载慢，应交同步登录再确认）；True+失败=密码错误或登录框
             卡住，应立即报错而不是继续跑。
    """

    def __init__(self) -> None:
        self.thread: threading.Thread | None = None
        self.outcome: bool | None = None
        self.saw_box = False
        self.error = ""

    def wait(self, timeout: float | None = None) -> bool | None:
        """等线程结束，返回 outcome（超时仍进行中返回 None）。"""
        if self.thread is not None:
            self.thread.join(timeout)
        return self.outcome

    @property
    def finished(self) -> bool:
        return self.outcome is not None


def _start_ao_login(flow_cfg: dict, expect_login: bool = True,
                    quiet: bool = False, timeout: float | None = None,
                    no_box_ok_after: float | None = None) -> LoginHandle | None:
    """后台线程自动登录 AO（返回结果句柄，未配置账号时返回 None）。

    AO 连接 BPC 时弹出的登录框是模态的，会阻塞 Excel 的 COM 消息泵，导致主线程的
    Workbooks 枚举卡住。这里用 uiautomation 在后台线程填登录框（不依赖 Excel COM），
    登录完成后模态框消失，主线程自然恢复。登录结果写入句柄的 outcome。

    quiet=True 用于「守候登录框」场景：等不到登录框属于正常情况，不告警。
    no_box_ok_after：登录框没出现时，等满这么多秒就按"已用缓存凭据登录"
    继续，不再干等到总体超时。仅用于「本工具刚启动 Excel」的场景。
    """
    try:
        from .orchestrator import login_config
        from .ao_login import AoLogin
    except Exception:
        return None
    login_cfg = login_config(flow_cfg)
    if not login_cfg.get("username"):
        return None
    wait = float(timeout if timeout is not None else login_cfg.get("timeout", 120))

    handle = LoginHandle()

    def _login() -> None:
        # uiautomation 需要在线程内初始化 COM（CoInitialize + 建 UIAutomation 对象），
        # 否则报 "尚未调用 CoInitialize"；用库自带的初始化器包一层。
        import uiautomation as auto
        with auto.UIAutomationInitializerInThread():
            try:
                ao = AoLogin(
                    username=login_cfg["username"],
                    password=login_cfg.get("password", ""),
                    client=str(login_cfg.get("client", "333")),
                    lang=str(login_cfg.get("lang", "zh")),
                )
                ok = ao.ensure(wait, expect_login=expect_login,
                               no_box_ok_after=no_box_ok_after)
                handle.outcome = bool(ok)
                handle.saw_box = ao.saw_login_box
                if not ok and not quiet:
                    log.warning("后台自动登录未成功（超时或密码错误），请检查账号密码")
                elif not ok and quiet and not handle.saw_box:
                    # quiet 只用于"复用 Excel 守候登录框"场景：没等到框多数是
                    # 已登录（正常），但也可能是 AO 加载慢、守候超时后框才弹。
                    # 打一条明确日志，事后好定位"卡死在登录框"是不是等短了。
                    log.info(
                        "复用 Excel：%.0fs 内未检测到 AO 登录框（已登录时属正常）；"
                        "若随后的操作卡死在登录框，请在 flow.yaml 调大 "
                        "login.reuse_wait_sec（当前 %.0fs）", wait, wait)
            except Exception as exc:
                handle.outcome = False
                handle.error = str(exc)
                if not quiet:
                    log.warning(f"后台自动登录异常: {exc}")

    t = threading.Thread(target=_login, daemon=True)
    handle.thread = t
    t.start()
    return handle


def _mark_login(engine: ExcelEngine, launched_now: bool, login_done: bool = False,
                login_handle: "LoginHandle | None" = None) -> ExcelEngine:
    """给引擎打上启动/登录状态，供 run_flow 判断是否还要再等登录框。

    launched_now:  本次是否由本函数启动了 Excel（AO 登录框可能还没弹出来）。
    login_done:    启动阶段已**确认**登录成功（或无需登录）——只有拿到登录
                   成功信号才能置 True；后台线程"已启动/已结束"不算数，
                   失败 outcome 由 run_flow 消费后给出明确报错。
    login_handle:  后台登录线程的结果句柄（未起线程时为 None）。
    login_settled: 登录后的 AO 收尾屏障（wait_ao_login_idle）是否已做过；
                   复用 Excel 的路径在 bootstrap 里只起线程不等结果，屏障
                   由 run_flow/recover 补，靠这个标志防止重复等待。
    """
    engine.launched_now = launched_now
    engine.login_done = login_done
    engine.login_handle = login_handle
    engine.login_settled = False
    return engine


def ensure_excel_ready(flow_cfg: dict, kill_stale: bool = False,
                       stale_pid: int | None = None, wait_sec: int = 180,
                       force_restart: bool = False):
    """确保真实 Excel 已运行且导航表已打开，否则自动打开。

    force_restart=True：AO 会话失效时进程和导航表都还在，复用旧实例无法
    重建会话，必须先结束本流程自己的 Excel 再重启。

    返回可复用的 ExcelEngine（所有成功路径），调用方可复用同一实例，避免重连到别的 Excel。
    打开导航表用真实 Excel 的 EXE 以文档参数启动（等价于双击文件），完整加载 AO 并连接
    BPC；若用 COM 的 Workbooks.Open，AO 常不初始化、进不了 PBF。
    """
    # 启动阶段也可能被取消（Web 界面「取消」）：这里没有检查点时，CLI 会一直
    # 等到任务超时，Web 端 60s 宽限后 taskkill /T 连刚打开的导航表 Excel 一起
    # 杀掉——现场表现就是「AO 又没了」。取消标志一出现就优雅退出，Excel/AO
    # 留在原地，用户可以接着手动操作。
    raise_if_cancelled()
    nav = flow_cfg.get("navigator") or {}
    pattern = nav.get("workbook_pattern", "PBF Pro - 全面预算系统流程导航*")
    path = Path(nav["launch_file"]) if nav.get("launch_file") else None

    # 0) 会话失效：强制结束本流程的 Excel。只认 stale_pid，绝不无差别杀进程。
    if force_restart:
        running = excel_pids()
        if stale_pid and stale_pid in running:
            # 先把导航表标记为已保存，避免 taskkill 后 Excel 下次启动弹恢复副本
            try:
                engine = ExcelEngine(timeout_sec=3, make_visible=False)
                for wb in engine._excel.Workbooks:
                    if "导航" in wb.Name:
                        wb.Saved = True
            except Exception:
                pass
            log.warning(f"AO 会话失效，结束旧 Excel PID={stale_pid} 以重建会话…")
            kill_excel(stale_pid)
        elif running:
            # 旧 PID 不在了但有其它 Excel 进程：要么旧 Excel 已死、要么 PID 记录
            # 丢了。Excel 是单实例，Dispatch 只会连上现有进程——导航表还开着就
            # 说明死会话的实例还活着，不结束它"重启"只是空转，后续宏照旧报
            # 0x800A9C68。宁可报错让用户手动关，也不悄悄复用死会话。
            has_nav = False
            try:
                probe = ExcelEngine(timeout_sec=3, make_visible=False)
                probe.find_workbook(pattern)
                has_nav = True
            except ExcelError:
                has_nav = False
            if has_nav:
                raise ExcelError(
                    f"AO 会话失效，但未能定位旧 Excel PID（目标 PID={stale_pid}），"
                    "而导航表仍开着——死会话的 Excel 还在运行。为避免误杀用户其它 "
                    "Excel 没有强制结束；请手动关闭该 Excel 后重试")
            log.warning(
                f"AO 会话失效，未能定位旧 Excel PID（目标 PID={stale_pid}）；"
                "导航表未开（旧实例可能已退出），按新会话重新启动")

    # 1) 已经就绪（真实 Excel 里导航表已打开）——仅当确有 Excel 进程时才去连，
    #    否则 ExcelEngine 的 Dispatch 会拉起一个空 Excel，还得再关掉，反而制造竞态。
    if excel_pids():
        engine = None
        has_path = path is not None and path.exists()
        if has_path:
            # 有多个 Excel 进程时（例如上一轮残留的无窗口实例），Dispatch 可能
            # 连到没有导航表的那个，find_workbook 会一直找不到。优先按本地文件
            # 路径从 ROT 精确绑到开着导航表的实例。
            try:
                engine = ExcelEngine(timeout_sec=3, workbook_path=str(path))
                engine.find_workbook(pattern)
            except ExcelError:
                engine = None
        if engine is None:
            # ROT 里没找到（或暂忙）：退回普通 Dispatch，正常单实例场景仍有效。
            # 探测不能改可见性——Dispatch 可能连上上一轮遗留的隐藏空实例，
            # 设成可见就会在用户屏幕上多弹一个空白 Excel；真找到导航表后下面
            # 会再 _ensure_visible。
            try:
                engine = ExcelEngine(timeout_sec=3, make_visible=False)
                engine.find_workbook(pattern)
            except ExcelError:
                engine = None
        if engine is not None:
            _ensure_visible(engine)
            # 这个实例在本次运行前就开着，属于「复用」——即便它是本工具上一轮
            # 启动的，也要靠已有 pbf 记录保留归属；不能把用户自己打开的 Excel
            # 误记成 pbf 实例而在以后自动关掉。
            _record_owner(engine, str(path or pattern),
                          launched_pids={engine.pid} if engine.pid else set())
            # 复用已有 Excel：用户可能刚打开导航表、登录框还没弹出来。
            # 起一个后台线程守候登录框（没出现就按超时自行退出），
            # 这样随后卡在模态登录框上的导航表/刷新调用会被自动解锁。
            # 结果句柄交给 run_flow 等 outcome，不能把"线程已启动"当登录成功。
            # 复用场景里"没有登录框"是正常情况（缓存凭据），expect_login=False。
            # 超时默认只给 15s：登录框 10s 左右就会弹，已登录时不会弹；
            # 可用 flow.yaml login.reuse_wait_sec 调整。
            handle = _start_ao_login(flow_cfg, expect_login=False, quiet=True,
                                     timeout=_reuse_login_wait_sec(flow_cfg))
            return _mark_login(engine, launched_now=False, login_handle=handle)

    # 2) 无窗口的残留实例：只有本工具启动过（owner 记录 source=pbf）的才自动关闭。
    #    用户/其它 COM 客户端打开的 Excel 一律不碰；Quit 后必须验证进程真的退出，
    #    否则新导航表会经单实例机制落回这个僵尸实例，AO 登录框弹不出来。
    if excel_pids():
        engine = None
        empty = False
        try:
            engine = ExcelEngine(timeout_sec=3, make_visible=False)
            empty = engine.workbook_count() == 0
        except ExcelError:
            engine = None
        if engine is not None and empty:
            owner = read_excel_owner()
            owned = bool(owner and owner.get("pid") == engine.pid
                         and owner.get("source") == OWNER_SOURCE_PBF)
            if owned:
                if not quit_excel_verified(engine, "清理本工具遗留的无窗口实例"):
                    raise ExcelError(
                        f"本工具遗留的 Excel 实例（PID={engine.pid}）无法结束："
                        "Quit 和强制结束都未生效。为避免在僵尸实例上继续操作，"
                        "流程已停止；请手动结束该 EXCEL.EXE 后重试")
                time.sleep(1)
            else:
                log.warning(
                    f"检测到无窗口 Excel PID={engine.pid}，但不是本工具启动的"
                    "（无归属记录）——不自动关闭，避免破坏其它会话或用户数据；"
                    "若流程卡住，请手动关闭该实例")

    # 3) 断开重连时：残留进程可能已经僵死，只结束本工具连接的那个 PID。
    #    定位不到目标时宁可跳过，也不要 taskkill 掉用户其它 Excel（可能有未保存内容）。
    if kill_stale:
        running = excel_pids()
        if stale_pid and stale_pid in running:
            kill_excel(stale_pid)
        elif running:
            log.warning(
                f"未定位到需要结束的 Excel 进程（目标 PID={stale_pid}），"
                f"跳过强制结束，避免影响其它 Excel 窗口")

    # 4) 打开导航表
    #    source=bw（默认）：启动空白 Excel → 文件 → Analysis → 从 BW 平台现取导航表。
    #      本地 .xlsb 会带缓存的连接/版本上下文，切版本容易出错；BW 每次都是新副本。
    #    source=file：直接打开 flow.yaml 里配置的本地文件（旧方式，保留兜底）。
    source = str(nav.get("source") or "bw").lower()
    if source == "bw":
        from .bw_open import open_navigator_from_bw

        # bw_open 内部自己 Popen 空白表；Excel 是单实例，也可能只是把启动请求
        # 交给已在运行的实例。启动前记基线，之后才能区分"新进程"和"复用"。
        launched_base = set(excel_pids())
        engine = open_navigator_from_bw(flow_cfg, wait_sec=wait_sec)
        _ensure_visible(engine)
        _record_owner(engine, str(path or pattern), launched_pids=launched_base)
        # bw_open 内部检测到登录框时会自动填写并等它关闭，走到这里说明登录已处理
        return _mark_login(engine, launched_now=True, login_done=True)

    if not path or not path.exists():
        raise ExcelError(
            f"导航表未打开，且路径无效：{path}；请检查 config/flow.yaml 的 navigator.launch_file")
    excel_exe = ExcelEngine._real_excel_path()
    if not excel_exe:
        raise ExcelError("找不到真实 Excel（EXCEL.EXE），无法打开导航表，请确认 Office 已安装")
    reset_active_excel()
    # 记录启动前的 Excel PID：Excel 是单实例，Popen 可能只是把文件交给已在运行的
    # 实例（新进程随即退出）。事后用基线区分"新进程"和"复用了别人的实例"。
    launched_base = set(excel_pids())
    opened_now = False
    if ExcelEngine.is_workbook_open(str(path)):
        # 第 1 步没绑上多半只是 AO 忙；文件其实已经开着，别再启动一份。
        log.info(f"导航表已在运行中的 Excel 里打开，等待 COM 可绑定：{path}")
        login_handle = _start_ao_login(
            flow_cfg, expect_login=False, quiet=True,
            timeout=_reuse_login_wait_sec(flow_cfg))
    else:
        log.info(f"正在打开导航表：{path}")
        try:
            subprocess.Popen([excel_exe, str(path)])
        except Exception as exc:
            raise ExcelError(f"启动 Excel 打开导航表失败：{exc}") from exc
        opened_now = True
        # 后台登录：模态登录框会阻塞 Excel COM，用 uiautomation 线程先把它填掉
        login_handle = _start_ao_login(
            flow_cfg, no_box_ok_after=launch_no_box_ok_sec(flow_cfg))

    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(3)
        raise_if_cancelled()
        try:
            # 按文件路径绑到真正开着导航表的实例；未出现/繁忙时抛错并继续轮询。
            engine = ExcelEngine(timeout_sec=5, workbook_path=str(path))
            engine.find_workbook(pattern)
            try:
                engine._excel.DisplayAlerts = False
            except Exception:
                pass
            log.info("导航表已打开，等待 AO 加载…")
            _ensure_visible(engine)
            _record_owner(engine, str(path), launched_pids=launched_base)
            time.sleep(10)
            raise_if_cancelled()
            if login_handle is not None:
                # 等登录出结果；失败不在这里吞掉——run_flow 看句柄 outcome
                # 决定跳过登录还是明确报错（密码错误时"线程已结束"≠登录成功）。
                # 分片等待，取消时能立即收尾，而不是拖到 Web 端硬杀进程树。
                login_deadline = time.time() + 90
                while time.time() < login_deadline:
                    raise_if_cancelled()
                    if login_handle.wait(timeout=1) is not None:
                        break
            # 登录框关闭后加载项会自动刷新所有数据源；它没刷完就切版本会把
            # 加载项打崩（现场 15:23 的「SAP BI 加载项已断开所有数据源连接」）。
            # 只在本次真的填过登录框时等，复用已有会话不额外拖延。
            # 必须先 _mark_login 再 settle：屏障完成要落在 login_settled=True
            # 上，先 settle 后 mark 会把标志清回 False，run_flow 会重复等一遍。
            mark = _mark_login(engine, launched_now=opened_now,
                               login_handle=login_handle)
            if (login_handle is not None and login_handle.outcome
                    and getattr(login_handle, "saw_box", False)):
                wait_ao_login_idle(mark, flow_cfg)
            return mark
        except ExcelError:
            continue
    raise ExcelError(f"等待导航表打开超时（{wait_sec}s）")
