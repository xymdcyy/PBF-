"""安全点取消：Web 界面写一个取消标志文件，引擎在安全点检查并优雅退出。

不走 taskkill 硬杀：硬杀会把 Excel 停在写入/保存中途，留下无窗口的僵尸实例，
下次运行连上去后 AO 登录框弹不出来，看起来就像"AO 没了"。

CLI 由 Web 界面启动时传入环境变量 PBF_CANCEL_FILE；命令行直接运行时没有这个
变量，取消检查永远是"未取消"，行为与以前一致。
"""
from __future__ import annotations

import os
from pathlib import Path

ENV_CANCEL_FILE = "PBF_CANCEL_FILE"

_cancel_file: Path | None = None


class CancelledError(RuntimeError):
    """用户在安全点请求取消。此时尚未保存的修改应被丢弃。"""


def set_cancel_file(path: str | Path | None) -> None:
    """设置取消标志文件路径；None/空串表示不启用取消检查。"""
    global _cancel_file
    raw = str(path or "").strip()
    _cancel_file = Path(raw) if raw else None


def current_cancel_file() -> Path | None:
    return _cancel_file


def set_cancel_file_from_env() -> None:
    set_cancel_file(os.environ.get(ENV_CANCEL_FILE))


def is_cancelled() -> bool:
    """取消标志文件存在即视为已请求取消。文件不存在/未配置 → False。"""
    path = _cancel_file
    if path is None:
        return False
    try:
        return path.exists()
    except OSError:
        return False


def raise_if_cancelled() -> None:
    """在安全点调用：已请求取消则抛 CancelledError，由上层收尾。"""
    if is_cancelled():
        raise CancelledError("已按用户请求取消")
