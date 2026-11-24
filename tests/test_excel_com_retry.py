"""ExcelEngine：AO 忙时的 RPC_E_CALL_REJECTED 必须重试，且不能泄漏裸 COM 异常。

现场（2026-11-24 14:03）：AO 刷新期间 workbook/cell 调用抛
(-2147418111, '被呼叫方拒绝接收呼叫。')，read/write 不重试，异常在
`raise ExcelError(f"... {sheet.Name} ...")` 里再次触发 COM 调用，最终以裸
com_error 冒到 Web，界面只看到一行「[错误] (-2147418111, ...)」。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _stubs  # noqa: F401

import engine.excel as excel_mod
from engine.excel import ExcelEngine, ExcelError


class _ComBusy(Exception):
    def __init__(self) -> None:
        super().__init__("(-2147418111, '被呼叫方拒绝接收呼叫。', None, None)")


class _Cell:
    def __init__(self, owner: "_Sheet") -> None:
        self.owner = owner

    @property
    def Value2(self):  # noqa: N802
        self.owner._maybe_reject()
        return self.owner.value

    @Value2.setter
    def Value2(self, value):  # noqa: N802
        self.owner._maybe_reject()
        self.owner.value = value

    def Select(self):  # noqa: N802
        self.owner._maybe_reject()


class _Sheet:
    def __init__(self, name="导航", value="OLD", busy=0, name_busy=False) -> None:
        self._name = name
        self.value = value
        self.busy = busy
        self.name_busy = name_busy

    @property
    def Name(self):  # noqa: N802
        if self.name_busy:
            raise _ComBusy()
        return self._name

    def _maybe_reject(self) -> None:
        if self.busy:
            self.busy -= 1
            raise _ComBusy()

    def Range(self, cell):  # noqa: N802
        return _Cell(self)

    def Activate(self):  # noqa: N802
        self._maybe_reject()

    def Select(self):  # noqa: N802
        self._maybe_reject()


class _Wb:
    def __init__(self, busy=0) -> None:
        self.busy = busy

    def Activate(self):  # noqa: N802
        if self.busy:
            self.busy -= 1
            raise _ComBusy()


class _Excel:
    def __init__(self, busy=0, result=1) -> None:
        self.busy = busy
        self.result = result
        self.calls = []

    def Run(self, *args):  # noqa: N802
        self.calls.append(args)
        if self.busy:
            self.busy -= 1
            raise _ComBusy()
        return self.result


def _engine():
    return object.__new__(ExcelEngine)


class ComRetryTest(unittest.TestCase):
    def setUp(self):
        patcher = mock.patch.object(excel_mod.time, "sleep")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_read_retries_rejected_calls(self):
        sheet = _Sheet(value="V", busy=2)
        self.assertEqual(_engine().read(sheet, "C20"), "V")

    def test_read_wraps_persistent_rejection(self):
        sheet = _Sheet(busy=99)
        with self.assertRaises(ExcelError) as ctx:
            _engine().read(sheet, "C20")
        self.assertIn("read 导航!C20", str(ctx.exception))
        self.assertIn("拒绝接收呼叫", str(ctx.exception))

    def test_read_does_not_leak_when_sheet_name_also_busy(self):
        # 以前 sheet.Name 写在 except 的 f-string 里，AO 忙时会二次抛裸 COM 异常
        sheet = _Sheet(busy=99, name_busy=True)
        with self.assertRaises(ExcelError) as ctx:
            _engine().read(sheet, "C20")
        self.assertIn("read ?!C20", str(ctx.exception))

    def test_write_retries_rejected_calls(self):
        sheet = _Sheet(busy=2)
        _engine().write(sheet, "C20", "KEY")
        self.assertEqual(sheet.value, "KEY")

    def test_select_retries_rejected_calls(self):
        sheet = _Sheet(busy=2)
        _engine().select(sheet, "C20")     # 不抛异常即通过

    def test_activate_retries_rejected_calls(self):
        _engine().activate(_Wb(busy=2))

    def test_ao_call_retries_and_reports_failure(self):
        excel = _Excel(busy=2)
        eng = _engine()
        eng._excel = excel
        self.assertEqual(eng.ao_call("SAPSetFilter", "DS_VAR", "ZDCATEGRY", "K"),
                         1)
        self.assertEqual(len(excel.calls), 3)

        eng._excel = _Excel(busy=99)
        with self.assertRaises(ExcelError) as ctx:
            eng.ao_call("SAPSetFilter")
        self.assertIn("AO SAPSetFilter", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
