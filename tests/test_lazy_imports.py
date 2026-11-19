"""engine 包的按需导入：纯子模块不应拉起 COM 依赖。

这些测试用子进程运行，避免当前解释器里已导入的模块（或测试桩）影响结果，
因此在没有 pywin32/uiautomation 的 CI 环境里也能通过。
"""
from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _run(code: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, "-c", code], cwd=ROOT,
                          capture_output=True, text=True, encoding="utf-8")


class TestLazyImports(unittest.TestCase):
    def test_import_engine_does_not_load_com(self):
        r = _run(
            "import sys\n"
            "import engine\n"
            "assert 'win32com.client' not in sys.modules\n"
            "assert 'pythoncom' not in sys.modules\n"
            "assert 'uiautomation' not in sys.modules\n"
        )
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_pure_submodule_import_does_not_load_com(self):
        r = _run(
            "import sys\n"
            "from engine import periods\n"
            "assert periods.parse_periods('2027.01-2027.03') == ['2027.01', '2027.02', '2027.03']\n"
            "assert 'win32com.client' not in sys.modules\n"
        )
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_export_table_consistent(self):
        import engine
        self.assertEqual(list(engine.__all__), sorted(engine._EXPORTS))
        for name in engine.__all__:
            self.assertIn(name, engine._EXPORTS)

    def test_export_names_exist_in_source(self):
        """静态检查 _EXPORTS 映射：每个名字都必须在其子模块源码中定义/导入。"""
        import ast

        import engine
        pkg = Path(engine.__file__).parent
        for name, mod in engine._EXPORTS.items():
            src = (pkg / (mod.lstrip(".") + ".py")).read_text(encoding="utf-8")
            tree = ast.parse(src)
            defined: set[str] = set()
            for node in tree.body:
                if isinstance(node, (ast.Import, ast.ImportFrom)):
                    defined |= {a.asname or a.name.split(".")[0] for a in node.names}
                elif isinstance(node, (ast.FunctionDef, ast.ClassDef)):
                    defined.add(node.name)
                elif isinstance(node, ast.Assign):
                    defined |= {t.id for t in node.targets if isinstance(t, ast.Name)}
            self.assertIn(name, defined, f"{mod} 中找不到导出名 {name}")

    def test_unknown_name_raises_attribute_error(self):
        import engine
        with self.assertRaises(AttributeError):
            engine.NoSuchThing

    def test_dir_contains_exports(self):
        import engine
        self.assertIn("ExcelEngine", dir(engine))


if __name__ == "__main__":
    unittest.main()
