"""业务单元（品牌 × 品类）配置校验与流程覆盖。

`config/business_units.yaml` 为每个业务单元按顺序给出各编制表的 `nav_cell`。
`apply_unit()` 把这些 `nav_cell` 覆盖到 `flow.yaml` 的 steps 上，让「品牌 × 品类」
选择真正决定打开导航表的哪一行（而不是永远跑 flow.yaml 里写死的乐华行）。

纯逻辑模块：不依赖 pywin32 / COM，也不直接读 YAML（由调用方传入已解析的 dict），
便于单元测试。
"""
from __future__ import annotations

import copy
import csv
import re
from pathlib import Path

DIM_COLUMNS = ("品牌", "品类")


class UnitError(ValueError):
    """业务单元配置错误（未知/未启用/步骤不对应等）。"""


def _norm(s: object) -> str:
    """表单名归一化：去掉空白和下划线后做包含匹配。"""
    return re.sub(r"[\s_]+", "", str(s or ""))


def units(units_cfg: dict) -> dict[str, dict]:
    return dict(units_cfg.get("units") or {})


def unit_names(units_cfg: dict) -> list[str]:
    return list(units(units_cfg))


def enabled_unit_names(units_cfg: dict) -> list[str]:
    return [name for name, u in units(units_cfg).items() if u.get("启用")]


def default_unit(units_cfg: dict) -> str | None:
    """默认业务单元：第一个启用的单元。"""
    names = enabled_unit_names(units_cfg)
    return names[0] if names else None


def get_unit(units_cfg: dict, name: str) -> dict:
    """取一个已启用的单元；未知或未启用抛 UnitError。"""
    all_units = units(units_cfg)
    if name not in all_units:
        raise UnitError(f"未知业务单元 '{name}'，可选：{list(all_units)}")
    unit = all_units[name] or {}
    if not unit.get("启用"):
        raise UnitError(
            f"业务单元 '{name}' 未启用（config/business_units.yaml 中 启用: false）")
    return unit


def _pair_steps(flow_steps: list[dict], unit: dict, name: str) -> list[tuple[dict, dict]]:
    """按顺序把单元 steps 与 flow steps 配对并校验。

    顺序是权威的（YAML 注释约定 steps 与 flow.yaml 顺序一一对应）；同时用
    `表单` 名称做一次包含匹配，防止配置顺序被改乱后静默覆盖到错误的表。
    """
    unit_steps = unit.get("steps") or []
    if len(unit_steps) != len(flow_steps):
        raise UnitError(
            f"业务单元 '{name}' 的 steps 数量 {len(unit_steps)} 与 "
            f"flow.yaml 的 {len(flow_steps)} 不一致")
    pairs: list[tuple[dict, dict]] = []
    for i, (flow_step, unit_step) in enumerate(zip(flow_steps, unit_steps), start=1):
        hint = _norm(unit_step.get("表单", ""))
        form = _norm(flow_step.get("form", ""))
        if hint and hint not in form:
            raise UnitError(
                f"业务单元 '{name}' 第 {i} 步表单 '{unit_step.get('表单')}' 与 "
                f"flow.yaml 的 '{flow_step.get('form')}' 不对应（steps 必须按顺序一一对应）")
        nav_cell = str(unit_step.get("nav_cell") or "").strip()
        if not nav_cell:
            raise UnitError(f"业务单元 '{name}' 第 {i} 步缺少 nav_cell")
        pairs.append((flow_step, unit_step))
    return pairs


def apply_unit(flow_cfg: dict, units_cfg: dict, name: str) -> dict:
    """返回覆盖了单元 nav_cell 的 flow 配置副本（不修改入参）。

    覆盖范围：
    - 每个 step 的 `nav_cell`（必填）；
    - step 可选的 `after_close_nav_cell` → `after_close.nav_cell`；
    - 单元可选的 `verify_nav_cell` / `verify_form_row` → `verify.*`。
    """
    unit = get_unit(units_cfg, name)
    flow = copy.deepcopy(flow_cfg)
    flow_steps = flow.get("steps") or []

    for flow_step, unit_step in _pair_steps(flow_steps, unit, name):
        flow_step["nav_cell"] = str(unit_step["nav_cell"]).strip()
        after = flow_step.get("after_close")
        if after and unit_step.get("after_close_nav_cell"):
            after["nav_cell"] = str(unit_step["after_close_nav_cell"]).strip()

    verify = flow.get("verify") or {}
    if unit.get("verify_nav_cell"):
        verify["nav_cell"] = str(unit["verify_nav_cell"]).strip()
        flow["verify"] = verify
    if unit.get("verify_form_row") is not None:
        verify["form_row"] = int(unit["verify_form_row"])
        flow["verify"] = verify

    flow["_unit"] = {
        "name": name,
        "brand": str(unit.get("品牌", "")),
        "category": str(unit.get("品类", "")),
    }
    return flow


def check_units(units_cfg: dict, flow_cfg: dict) -> list[str]:
    """离线校验业务单元配置，返回人类可读的问题列表（空列表 = 通过）。

    级别前缀：`错误:` 会阻止运行，`告警:` 仅提示。
    """
    issues: list[str] = []
    brands = units_cfg.get("品牌") or []
    categories = units_cfg.get("品类") or []
    all_units = units(units_cfg)

    if not all_units:
        issues.append("错误: 未配置任何业务单元")
        return issues
    if not enabled_unit_names(units_cfg):
        issues.append("错误: 没有启用的业务单元（至少一个单元 启用: true）")

    for name, unit in all_units.items():
        unit = unit or {}
        brand, category = str(unit.get("品牌", "")), str(unit.get("品类", ""))
        if not brand or not category:
            issues.append(f"错误: 单元 '{name}' 缺少 品牌/品类")
        else:
            if brands and brand not in brands:
                issues.append(f"告警: 单元 '{name}' 的品牌 '{brand}' 不在 品牌 列表中")
            if categories and category not in categories:
                issues.append(f"告警: 单元 '{name}' 的品类 '{category}' 不在 品类 列表中")

        if not unit.get("启用"):
            continue
        try:
            pairs = _pair_steps(flow_cfg.get("steps") or [], unit, name)
        except UnitError as exc:
            issues.append(f"错误: {exc}")
            continue
        cells = [str(us["nav_cell"]).strip() for _, us in pairs]
        dup = sorted({c for c in cells if cells.count(c) > 1})
        if dup:
            issues.append(f"错误: 单元 '{name}' 的 nav_cell 重复：{dup}")
    return issues


def _read_csv(path: Path) -> list[list[str]]:
    for encoding in ("utf-8-sig", "gbk"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                return list(csv.reader(f))
        except UnicodeDecodeError:
            continue
    raise OSError(f"无法识别 CSV 编码: {path}")


def csv_dimension_warnings(out_dir: Path | str, dims: tuple[str, ...] = DIM_COLUMNS) -> list[str]:
    """检查导出目录里的 CSV：品牌/品类 列存在但整列为空时给出告警。

    「品牌品类为空」意味着表单的组合键匹配会整行 skipped，提前提示比跑到
    Excel 里才发现更省时间。
    """
    warnings: list[str] = []
    for path in sorted(Path(out_dir).glob("*.csv")):
        try:
            rows = _read_csv(path)
        except OSError as exc:
            warnings.append(f"{path.name}: 读取失败（{exc}）")
            continue
        if len(rows) < 2:
            continue
        header = [str(h).strip() for h in rows[0]]
        for dim in dims:
            if dim not in header:
                continue
            idx = header.index(dim)
            values = [row[idx] for row in rows[1:] if idx < len(row)]
            if values and not any(str(v).strip() for v in values):
                warnings.append(f"{path.name}: '{dim}' 列全部为空（请检查模板表对应列）")
    return warnings
