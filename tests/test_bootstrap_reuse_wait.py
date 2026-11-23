"""login.reuse_wait_sec 配置解析：非法值告警并回退默认，不悄悄改变行为。"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine.bootstrap import REUSE_LOGIN_WAIT_SEC, _reuse_login_wait_sec


class ReuseWaitSecTests(unittest.TestCase):
    def test_default_when_absent(self) -> None:
        # 没配 / login 段为空 / login 段不是映射 → 都走默认值
        self.assertEqual(_reuse_login_wait_sec({}), REUSE_LOGIN_WAIT_SEC)
        self.assertEqual(_reuse_login_wait_sec({"login": None}), REUSE_LOGIN_WAIT_SEC)
        self.assertEqual(_reuse_login_wait_sec({"login": "不是映射"}),
                         REUSE_LOGIN_WAIT_SEC)

    def test_valid_value(self) -> None:
        self.assertEqual(_reuse_login_wait_sec({"login": {"reuse_wait_sec": 60}}), 60.0)
        self.assertEqual(_reuse_login_wait_sec({"login": {"reuse_wait_sec": "30"}}),
                         30.0)  # yaml 里引号包着的数字也接受

    def test_invalid_value_falls_back(self) -> None:
        # True/False 是 int 子类：True 会 float 成 1.0，须显式排除
        for bad in ("abc", object(), True, False, 0, -5):
            with self.subTest(bad=bad):
                self.assertEqual(_reuse_login_wait_sec({"login": {"reuse_wait_sec": bad}}),
                                 REUSE_LOGIN_WAIT_SEC)


if __name__ == "__main__":
    unittest.main()
