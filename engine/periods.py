"""期间解析工具 — 把用户输入的期间描述解析成有序期间列表，并提供期间序比较。

用法:
    from engine.periods import parse_periods, period_index
    parse_periods("2027.01")                       # -> ["2027.01"]
    parse_periods("2027.01,2027.03")               # -> ["2027.01", "2027.03"]
    parse_periods("2027.01-2027.04")               # -> ["2027.01".."2027.04"]
    parse_periods("2027.01,2027.03-2027.05")       # -> 混合，升序去重
"""
from __future__ import annotations

import re


def period_index(period: str) -> int | None:
    """'2027.03' → 月序号（便于比较先后）；格式非法返回 None。"""
    m = re.fullmatch(r"(\d{4})\.(\d{2})", (period or "").strip())
    return int(m.group(1)) * 12 + int(m.group(2)) - 1 if m else None


def _norm(p: str) -> str:
    """规范化单个期间为 'YYYY.MM'（允许 '2027.1' → '2027.01'），格式非法抛 ValueError。"""
    m = re.fullmatch(r"(\d{4})\.(\d{1,2})", (p or "").strip())
    if not m or not 1 <= int(m.group(2)) <= 12:
        raise ValueError(f"期间格式应为 YYYY.MM，实际为 {p!r}")
    return f"{m.group(1)}.{int(m.group(2)):02d}"


def parse_periods(spec: str) -> list[str]:
    """解析期间描述，返回升序去重列表。

    支持单值、逗号列表、`a-b` 区间，及混合（如 "2027.01,2027.03-2027.05"）。
    """
    if not spec or not str(spec).strip():
        raise ValueError("期间描述为空")
    tokens = [t.strip() for t in str(spec).split(",") if t.strip()]
    out: set[str] = set()
    for tok in tokens:
        if "-" in tok:
            # 区间：a-b
            a, b = tok.split("-", 1)
            ia, ib = period_index(_norm(a.strip())), period_index(_norm(b.strip()))
            if ia is None or ib is None:
                raise ValueError(f"期间区间非法: {tok}")
            if ia > ib:
                ia, ib = ib, ia
            for idx in range(ia, ib + 1):
                y, m = idx // 12, idx % 12 + 1
                out.add(f"{y}.{m:02d}")
        else:
            out.add(_norm(tok))
    return sorted(out, key=period_index)
