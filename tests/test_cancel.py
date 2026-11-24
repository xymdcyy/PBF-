"""安全点取消：标志文件、异常与写入前的中止（不碰 Excel/COM）。"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

from engine import cancel
from engine.cancel import CancelledError
from engine.writer import FormWriter, PreflightResult


class CancelFlagTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "run.cancel"
        cancel.set_cancel_file(self.path)

    def tearDown(self):
        cancel.set_cancel_file(None)
        self._tmp.cleanup()

    def test_no_flag_not_cancelled(self):
        self.assertFalse(cancel.is_cancelled())
        cancel.raise_if_cancelled()      # 不应抛异常

    def test_flag_file_means_cancelled(self):
        self.path.write_text("", encoding="utf-8")
        self.assertTrue(cancel.is_cancelled())
        with self.assertRaises(CancelledError):
            cancel.raise_if_cancelled()

    def test_flag_removed_is_not_cancelled(self):
        self.path.write_text("", encoding="utf-8")
        self.path.unlink()
        self.assertFalse(cancel.is_cancelled())

    def test_none_disables_check_even_if_file_exists(self):
        self.path.write_text("", encoding="utf-8")
        cancel.set_cancel_file(None)
        self.assertFalse(cancel.is_cancelled())

    def test_env_helper_sets_path(self):
        with mock.patch.dict(os.environ, {cancel.ENV_CANCEL_FILE: str(self.path)}):
            cancel.set_cancel_file(None)
            cancel.set_cancel_file_from_env()
        self.assertEqual(cancel.current_cancel_file(), self.path)

    def test_env_helper_empty_disables(self):
        cancel.set_cancel_file(self.path)
        with mock.patch.dict(os.environ, {cancel.ENV_CANCEL_FILE: ""}):
            cancel.set_cancel_file_from_env()
        self.assertIsNone(cancel.current_cancel_file())


class WriterCancelTest(unittest.TestCase):
    def tearDown(self):
        cancel.set_cancel_file(None)

    def test_write_and_save_stops_before_touching_cells(self):
        """取消标志存在时，write_and_save 不解析工作簿、不写任何单元格。"""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "run.cancel"
            path.write_text("", encoding="utf-8")
            cancel.set_cancel_file(path)
            writer = FormWriter(engine=object(), form_config={})
            with self.assertRaises(CancelledError):
                writer.write_and_save(PreflightResult(ok=True))

    def test_preflight_failure_still_raises_excel_error_first(self):
        from engine.excel import ExcelError

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "run.cancel"
            path.write_text("", encoding="utf-8")
            cancel.set_cancel_file(path)
            writer = FormWriter(engine=object(), form_config={})
            with self.assertRaises(ExcelError):
                writer.write_and_save(PreflightResult(ok=False))


if __name__ == "__main__":
    unittest.main()
