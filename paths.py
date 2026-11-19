"""运行目录解析 — 兼容源码运行和 PyInstaller 打包后的 exe。

打包成 exe 后 __file__ 指向临时解包目录，配置/日志/数据目录必须取 exe 所在目录。
"""
from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    """程序根目录：exe 所在目录；源码运行时是项目目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
