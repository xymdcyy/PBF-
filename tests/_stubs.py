"""纯逻辑测试的导入辅助。

测试只覆盖不碰 Excel/COM 的逻辑（期间解析、目录选择、CSV 导出/清除）。
在没装 pywin32 / uiautomation 的环境里，用占位模块让 engine 包可以导入；
这些占位模块不参与任何实际调用。
"""
from __future__ import annotations

import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

for _name in (
    "win32com",
    "win32com.client",
    "win32com.server",
    "win32com.server.exception",
    "uiautomation",
    "pythoncom",
):
    if _name in sys.modules:
        continue
    try:
        __import__(_name)
    except ImportError:
        sys.modules[_name] = types.ModuleType(_name)
