"""配置路径解析：机器相关的路径用环境变量，别写死在 flow.yaml 里。"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401  （确保仓库根在 sys.path）

import paths


class ResolvePathTest(unittest.TestCase):
    def test_env_var_expanded(self):
        with mock.patch.dict(os.environ, {"PBF_TEST_NAV_DIR": "/tmp/pbf-test"}):
            self.assertEqual(paths.resolve_path("${PBF_TEST_NAV_DIR}/nav.xlsb"),
                             Path("/tmp/pbf-test/nav.xlsb"))

    def test_relative_path_resolved_against_base(self):
        base = Path("/tmp/pbf-base")
        self.assertEqual(paths.resolve_path("nav.xlsb", base), base / "nav.xlsb")

    def test_absolute_path_kept(self):
        self.assertEqual(paths.resolve_path("/opt/nav.xlsb"), Path("/opt/nav.xlsb"))

    def test_expanduser(self):
        resolved = str(paths.resolve_path("~/nav.xlsb"))
        self.assertTrue(resolved.startswith(str(Path.home())))

    def test_unresolved_var_detected(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertTrue(paths.has_unresolved_var("${PBF_NOT_SET}/nav.xlsb"))
            self.assertFalse(paths.has_unresolved_var("/opt/nav.xlsb"))
