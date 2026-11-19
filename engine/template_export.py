"""从《滚动预测模板表.xlsx》导出上传数据（openpyxl 版，无需打开 Excel）。

支持单期间与多期间（1–6 个月）导出，期间从模板表头动态识别。
用法:
    from engine.template_export import export_template
    export_template(Path("滚动预测模板表.xlsx"), ["2027.03"], Path("data_2027.03"))
    export_template(Path("滚动预测模板表.xlsx"), "2027.01-2027.04", Path("data_2027.01_2027.04"))
    export_template(Path("滚动预测模板表.xlsx"), ["2027.01"], Path("data_clear"), zero=True)
"""
from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .periods import parse_periods, period_index


def _s(v: Any) -> str:
    """转字符串并去首尾空白；None → 空串。"""
    return "" if v is None else str(v).strip()


def _rows(ws) -> list[tuple]:
    return list(ws.iter_rows(values_only=True))


def _is_period(v: Any) -> bool:
    return isinstance(v, str) and re.fullmatch(r"\d{4}\.\d{2}", v.strip())


def _to_num(v: Any) -> float | None:
    """转数值（空/非数值 → None，用于求和聚合）。"""
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _normalize_periods(periods: Any) -> list[str]:
    """把 periods（str 描述或列表）规范化成升序去重的 'YYYY.MM' 列表（1–6 个）。"""
    if isinstance(periods, str):
        periods = parse_periods(periods)
    out: list[str] = []
    for p in periods:
        p = _s(p)
        if not re.fullmatch(r"\d{4}\.\d{2}", p):
            raise ValueError(f"期间格式应为 YYYY.MM，实际为 {p!r}")
        if p not in out:
            out.append(p)
    out = sorted(out, key=period_index)
    if not 1 <= len(out) <= 6:
        raise ValueError(f"期间数量应为 1–6，实际 {len(out)}")
    return out


def _dump(out_dir: Path, name: str, rows: list[dict]) -> int:
    if not rows:
        return 0
    path = out_dir / name
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    return len(rows)


def export_template(xlsx_path: Path, periods: Any, out_dir: Path, zero: bool = False) -> dict[str, int]:
    """导出 5 份 CSV 到 out_dir，返回 {文件名: 行数}。

    periods 可为 'YYYY.MM'、'YYYY.MM-YYYY.MM'、'YYYY.MM,YYYY.MM' 或列表。
    zero=True 时所有值填 0（供批量清除复用）。
    公式未缓存（data_only 读不到值）时自动回退 data_only=False 重读。
    """
    periods = _normalize_periods(periods)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"模板文件不存在: {xlsx_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    def load(data_only: bool):
        return load_workbook(xlsx_path, read_only=True, data_only=data_only)

    wb = load(True)
    result = _export(wb, periods, out_dir, zero)
    if sum(result.values()) == 0:
        wb.close()
        wb = load(False)
        result = _export(wb, periods, out_dir, zero)
    wb.close()
    return result


def _export(wb, periods: list[str], out_dir: Path, zero: bool) -> dict[str, int]:
    result: dict[str, int] = {}

    result["价格销量CT.csv"] = _dump(
        out_dir, "价格销量CT.csv",
        _export_measures(wb["价格&销量&加权CT"], periods, zero))
    result["客户费用.csv"] = _dump(
        out_dir, "客户费用.csv",
        _export_wide(wb["客户费用"], header_row=1, dim_cols=[0, 1],
                     dim_names=["渠道", "科目"], data_start=2, periods=periods, zero=zero))
    result["HQ毛上调整.csv"] = _dump(
        out_dir, "HQ毛上调整.csv", _export_hq(wb["HQ毛上调整"], periods, zero))
    result["其他成本.csv"] = _dump(
        out_dir, "其他成本.csv",
        _export_wide(wb["其他成本"], header_row=1, dim_cols=[2],
                     dim_names=["item"], data_start=2, periods=periods, zero=zero))
    result["费用及损益其他项.csv"] = _dump(
        out_dir, "费用及损益其他项.csv",
        _export_wide(wb["变动费用"], header_row=1, dim_cols=[0, 1, 2, 3, 4, 5],
                     dim_names=["组织", "科目", "产品", "区域", "渠道", "品牌"],
                     data_start=2, periods=periods, zero=zero))
    return result


# ---------- 宽表（simple / keyed 布局，期间在列上） ----------

def _export_wide(ws, header_row: int, dim_cols: list[int], dim_names: list[str],
                 data_start: int, periods: list[str], zero: bool) -> list[dict]:
    """宽表：维度列在 dim_cols，期间列在 header_row 表头行里动态识别，按维度键求和聚合。"""
    rows = _rows(ws)
    header = rows[header_row]
    per_col: dict[str, int] = {}
    for i, v in enumerate(header):
        if _is_period(v) and _s(v) not in per_col:
            per_col[_s(v)] = i

    agg: dict[tuple, dict[str, float]] = {}
    order: list[tuple] = []
    for r in rows[data_start:]:
        dims = tuple(_s(r[c]) for c in dim_cols)
        if not any(dims):
            continue
        if dims not in agg:
            agg[dims] = {}
            order.append(dims)
        for p in periods:
            if p not in per_col:
                continue
            if zero:
                agg[dims][p] = 0
                continue
            v = _to_num(r[per_col[p]])
            if v is not None:
                agg[dims][p] = agg[dims].get(p, 0.0) + v

    out: list[dict] = []
    for dims in order:
        row = {name: dims[i] for i, name in enumerate(dim_names)}
        has_any = False
        for p in periods:
            if zero:
                row[p] = 0
                has_any = True
            elif p in agg[dims]:
                row[p] = agg[dims][p]
                has_any = True
            else:
                row[p] = ""                      # 空期间写空串，填表时按空跳过
        if has_any:
            out.append(row)
    return out


# ---------- HQ毛上调整（matrix 布局，区域名取自行 2 的 B 列） ----------

def _export_hq(ws, periods: list[str], zero: bool) -> list[dict]:
    rows = _rows(ws)
    header = rows[0]
    per_col: dict[str, int] = {}
    for i, v in enumerate(header):
        if _is_period(v) and _s(v) not in per_col:
            per_col[_s(v)] = i
    region = _s(rows[2][1]) if len(rows) > 2 and len(rows[2]) > 1 else ""

    agg: dict[str, dict[str, float]] = {}
    order: list[str] = []
    for r in rows[3:]:
        a = _s(r[0]) if len(r) > 0 else ""
        if not a:
            continue
        if a not in agg:
            agg[a] = {}
            order.append(a)
        for p in periods:
            if p not in per_col:
                continue
            if zero:
                agg[a][p] = 0
                continue
            v = _to_num(r[per_col[p]])
            if v is not None:
                agg[a][p] = agg[a].get(p, 0.0) + v

    out: list[dict] = []
    for a in order:
        row = {"科目": a, "区域": region}
        has_any = False
        for p in periods:
            if zero:
                row[p] = 0
                has_any = True
            elif p in agg[a]:
                row[p] = agg[a][p]
                has_any = True
            else:
                row[p] = ""
        if has_any:
            out.append(row)
    return out


# ---------- 价格&销量&加权CT（measure_blocks 布局，期间在行上） ----------

def _export_measures(ws, periods: list[str], zero: bool) -> list[dict]:
    """长表：每个 (产品, 期间) 一行，字段 组织/渠道/品牌/产品/期间 + 3 个度量值。"""
    rows = _rows(ws)
    row0 = rows[0]                 # 度量组名（单价/销量/加权CT 各 6 列）
    row1 = rows[1]                 # 维度列 + 3×6 期间列

    per_order: list[str] = []
    for v in row1[4:]:
        if _is_period(v) and _s(v) not in per_order:
            per_order.append(_s(v))

    groups: list[str] = []
    for v in row0[4:]:
        label = _s(v).replace("\n", " ").replace("（台）", "").strip()
        if label and (not groups or groups[-1] != label):
            groups.append(label)

    n_meas = len(groups)
    n_per = len(per_order)
    block_starts = [4 + k * n_per for k in range(n_meas)]
    per_pos = {p: i for i, p in enumerate(per_order)}

    out: list[dict] = []
    for r in rows[2:]:
        dims = [_s(r[0]), _s(r[1]), _s(r[2]), _s(r[3])]
        if not any(dims):
            continue
        for p in periods:
            if p not in per_pos:
                continue
            idx = per_pos[p]
            vals = [r[block_starts[k] + idx] for k in range(n_meas)]
            if zero:
                measures = {groups[k]: 0 for k in range(n_meas)}
            else:
                if all(v in (None, "") for v in vals):
                    continue
                measures = {groups[k]: ("" if v in (None, "") else v) for k, v in enumerate(vals)}
            out.append({"组织": dims[0], "渠道": dims[1], "品牌": dims[2], "产品": dims[3],
                        "期间": p, **measures})
    return out
