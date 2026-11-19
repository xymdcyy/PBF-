"""SendKeys 文本转义（.NET SendKeys / uiautomation 语法）。

uiautomation 的 SendKeys 会把 + ^ % ~ ( ) { } 当成控制字符：
密码里出现这些字符时会被吞掉或触发组合键（如 % 触发 Alt 组合），
导致实际输入与配置不一致、登录反复失败。这里逐字符转义成字面量。
"""
from __future__ import annotations

_SPECIAL = set("+^%~()")


def escape_sendkeys(text: str) -> str:
    """把任意文本转成 SendKeys 可按字面发送的字符串。"""
    out: list[str] = []
    for ch in text:
        if ch in _SPECIAL:
            out.append("{" + ch + "}")
        elif ch == "{":
            out.append("{{}")
        elif ch == "}":
            out.append("{}}")
        else:
            out.append(ch)
    return "".join(out)
