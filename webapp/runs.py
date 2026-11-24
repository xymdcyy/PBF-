"""运行注册表：同一时刻只允许一个流程，并保留最近若干次运行状态。

独立于 Flask / engine，方便单元测试；线程安全。
"""
from __future__ import annotations

import threading

MAX_KEPT_RUNS = 20


class RunRegistry:
    """进程内运行状态表。

    begin() 原子地抢占「当前运行」槽位；end() 释放。只有持有槽位的运行才能
    被注册，从而避免两个请求同时通过「是否在运行」检查后再各自启动子进程。
    """

    def __init__(self, max_kept: int = MAX_KEPT_RUNS):
        self._runs: dict[str, dict] = {}
        self._lock = threading.RLock()
        self._current: str | None = None
        self._max_kept = max_kept

    def begin(self, run_id: str, state: dict) -> bool:
        """注册一次运行；已有运行在进行中时返回 False（不注册）。"""
        with self._lock:
            if self._current is not None:
                return False
            self._current = run_id
            self._runs[run_id] = state
            self._prune_locked()
            return True

    def end(self, run_id: str) -> None:
        with self._lock:
            if self._current == run_id:
                self._current = None

    def get(self, run_id: str) -> dict | None:
        with self._lock:
            return self._runs.get(run_id)

    @property
    def current(self) -> str | None:
        with self._lock:
            return self._current

    def _prune_locked(self) -> None:
        """只淘汰已结束的最旧运行，保持内存与日志引用不无限增长。"""
        if len(self._runs) <= self._max_kept:
            return
        finished = [rid for rid, s in self._runs.items()
                    if s.get("done") and rid != self._current]
        for rid in finished[:len(self._runs) - self._max_kept]:
            self._runs.pop(rid, None)


def run_owned_pids(state: dict, current_pids) -> set[int]:
    """本次运行"拥有"的 Excel PID：运行开始时不在、结束时还在的那些。

    之前只要有任何 EXCEL.EXE 就占槽，把用户自己开的 Excel（甚至别的 COM
    客户端开的）也算进来，逼着用户关掉唯一带 AO 会话的实例，反而毁掉会话。
    归属记录里的 pbf 实例（跨进程持久化）也一并计入。
    """
    if current_pids is None:
        return set()
    current = {int(p) for p in current_pids}
    baseline = {int(p) for p in (state.get("excel_pids_before") or [])}
    owned = {int(p) for p in (state.get("excel_owned_pids") or [])}
    owned |= current - baseline
    return current & owned


def should_hold_slot(state: dict, proc_alive: bool, current_pids=None) -> bool:
    """取消/超时后是否继续占用运行槽位。

    占位的目的：子进程没被确认终止，或**本次运行启动的** Excel 可能还在执行
    宏（COM 调用在服务端继续跑，taskkill /T 杀不到），此时放行新流程会出现
    两个流程操作同一个 Excel。

    - 子进程还活着 → 占位。
    - 正常结束 → 放行。
    - 取消后子进程已在安全点自行退出（cancel_graceful）→ 放行：此时 Excel
      可能还开着导航表，下一轮直接复用即可。
    - 硬取消/超时 → 只等本次运行启动的 Excel 退出；用户原有的 Excel 不占槽。
    """
    if proc_alive:
        return True
    if not (state.get("timed_out") or state.get("cancelled")):
        return False
    if state.get("cancel_package_running"):
        # 运行包已在 BPC 后台触发，取消只让工具不再等待；它可能仍在执行。
        # 无法从本地探测执行进度，保守占位，由用户确认后「强制释放槽位」。
        return True
    if state.get("cancel_graceful"):
        return False
    if current_pids is None:
        # 无法枚举进程：宁可多占一会儿，也不能让两个流程同时操作 Excel
        return True
    return bool(run_owned_pids(state, current_pids))


def can_force_release(state: dict, proc_alive: bool) -> bool:
    """是否允许人工强制释放槽位。

    只放行"运行已结束、被标记 stuck、子进程确已退出"的状态；仍在运行的
    子进程一律拒绝，避免人工操作放行后两个流程同时操作 Excel。
    用户点击前需自行确认 Excel 中没有仍在执行的宏。
    """
    return bool(state.get("done") and state.get("stuck")) and not proc_alive
