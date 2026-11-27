"""运行目录解析 — 兼容源码运行和 PyInstaller 打包后的 exe。

打包成 exe 后 __file__ 指向临时解包目录，配置/日志/数据目录必须取 exe 所在目录。
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

# 展开后仍形如 $VAR / ${VAR} 的残留，说明环境变量没设置
_UNRESOLVED_VAR = re.compile(r"\$[A-Za-z_{]")


def app_root() -> Path:
    """程序根目录：exe 所在目录；源码运行时是项目目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_path(value: str | Path, base: Path | None = None) -> Path:
    """把配置里的路径解析成绝对路径。

    支持 `${VAR}` / `$VAR`（Windows 下也认 `%VAR%`）与 `~`，相对路径按 base
    （默认程序根目录）展开。flow.yaml 里的 `launch_file` 这类**跟机器绑定**的
    路径一律写成环境变量，避免 Windows/mac 两台机器各改一份配置。
    """
    text = os.path.expanduser(os.path.expandvars(str(value)))
    path = Path(text)
    if not path.is_absolute():
        path = (base or app_root()) / path
    return path


def has_unresolved_var(value: str | Path) -> bool:
    """展开后仍残留变量占位（环境变量没设置）。"""
    return bool(_UNRESOLVED_VAR.search(os.path.expandvars(str(value))))
