"""从《滚动预测模板表.xlsx》导出上传数据（薄封装，委托 engine.template_export）。

用法:
    python export_templates.py 2027.03                              # 单期间（快捷）
    python export_templates.py --periods 2027.01-2027.04            # 多期间
    python export_templates.py --periods 2027.01,2027.03-2027.05    # 混合
    python export_templates.py --periods 2027.01-2027.04 --template 路径 --zero

等价于 `python cli.py export ...`，此处保留是为了向后兼容旧的
`python export_templates.py 2027.03` 用法。
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from engine.periods import parse_periods
from engine.template_export import export_template
from paths import app_root

ROOT = app_root()


def _default_template() -> Path:
    # 模板位置跟机器绑定：优先环境变量，其次项目根，最后桌面上的历史位置
    candidates = [
        Path(os.environ["PBF_TEMPLATE"]) if os.environ.get("PBF_TEMPLATE") else None,
        ROOT / "滚动预测模板表.xlsx",
        Path.home() / "Desktop" / "AR项目" / "滚动预测模板表.xlsx",
    ]
    for c in candidates:
        if c and c.exists():
            return c
    return ROOT / "滚动预测模板表.xlsx"


def _default_out_dir(periods: list[str]) -> Path:
    first, last = periods[0], periods[-1]
    name = f"data_{first}" if first == last else f"data_{first}_{last}"
    return ROOT / name


def main() -> None:
    p = argparse.ArgumentParser(description="从滚动预测模板表导出上传数据")
    p.add_argument("period_arg", nargs="?", help="单个期间，如 2027.03（快捷方式）")
    p.add_argument("--periods", help="期间，支持单值/列表/区间")
    p.add_argument("--template", help="模板 .xlsx 路径，默认项目根目录滚动预测模板表.xlsx")
    p.add_argument("--out-dir", help="输出目录，默认 data_<首>_<末>")
    p.add_argument("--zero", action="store_true", help="所有数值填 0")
    args = p.parse_args()

    spec = args.periods or args.period_arg
    if not spec:
        p.error("需提供期间：位置参数或 --periods")
    periods = parse_periods(spec)
    template = Path(args.template) if args.template else _default_template()
    out_dir = Path(args.out_dir) if args.out_dir else _default_out_dir(periods)

    print(f"模板: {template}")
    print(f"期间: {', '.join(periods)}")
    print(f"输出: {out_dir}")
    rows = export_template(template, periods, out_dir, zero=args.zero)
    for name, n in rows.items():
        print(f"  {name}: {n} 行")


if __name__ == "__main__":
    sys.exit(main())
