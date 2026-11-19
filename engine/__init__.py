"""PBF Uploader engine — Excel COM layer + writer + runner + navigator.

对外导出的名字按需导入（PEP 562）：`import engine` 或 `from engine import periods`
不会连带导入 pywin32/uiautomation，因此纯逻辑子模块（periods、header_periods、
config_check 等）可以在没装 Office/COM 的环境里直接测试。
"""
from __future__ import annotations

from importlib import import_module

_EXPORTS = {
    "ExcelEngine": ".excel",
    "ExcelError": ".excel",
    "is_disconnect": ".excel",
    "ensure_excel_ready": ".bootstrap",
    "kill_excel": ".bootstrap",
    "FormWriter": ".writer",
    "PreflightResult": ".writer",
    "RunPackageRunner": ".runner",
    "RunResult": ".runner",
    "FormReader": ".reader",
    "Navigator": ".navigator",
    "FormScanner": ".scanner",
    "PbfFlow": ".orchestrator",
    "run_flow": ".orchestrator",
    "AoLogin": ".ao_login",
}

__all__ = sorted(_EXPORTS)


def __getattr__(name: str):
    """首次访问时导入对应子模块并缓存结果（之后的访问走模块字典）。"""
    module_name = _EXPORTS.get(name)
    if module_name is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module = import_module(module_name, __name__)
    value = getattr(module, name)
    globals()[name] = value
    return value


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
