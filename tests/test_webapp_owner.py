"""Web 槽位：跨进程 Excel 归属记录的并入口径——只并入本轮运行期间写过的记录。

上一轮正常结束后 Excel 一直开着（reused 的常态），新一轮在触碰 Excel 之前
就被取消/超时硬杀时，旧记录若被并入，槽位会被与本次运行无关的 Excel 占住，
直到用户关 Excel 或手工强制释放才放行。
"""
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:
    import webapp.app as wa
    HAS_FLASK = True
except ImportError:                     # 轻量环境未装 flask：跳过，CI 装了会跑
    wa = None
    HAS_FLASK = False


@unittest.skipUnless(HAS_FLASK, "需要 flask（pip install -r requirements.txt）")
class ExcelOwnerPidTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.owner_file = Path(self._tmp.name) / "excel_owner.json"

    def _set_record(self, mtime: float | None = None) -> None:
        self.owner_file.write_text(
            '{"pid": 4242, "source": "reused"}', encoding="utf-8")
        if mtime is not None:
            os.utime(self.owner_file, (mtime, mtime))

    def _patch_engine(self, record):
        return (mock.patch("engine.bootstrap.owner_file",
                           return_value=self.owner_file),
                mock.patch("engine.bootstrap.read_excel_owner",
                           return_value=record))

    def test_record_from_this_run_counts(self):
        started = time.time() - 5
        self._set_record(mtime=time.time())          # 本轮运行期间写的
        f1, f2 = self._patch_engine({"pid": 4242, "source": "reused"})
        with f1, f2:
            self.assertEqual(wa._excel_owner_pid(started), 4242)

    def test_stale_record_from_previous_run_ignored(self):
        started = time.time()
        self._set_record(mtime=started - 3600)       # 上一轮写的
        f1, f2 = self._patch_engine({"pid": 4242, "source": "reused"})
        with f1, f2:
            self.assertIsNone(wa._excel_owner_pid(started))

    def test_without_scoping_still_counts(self):
        # 不传运行开始时间（无上下文）保持旧行为：记录有效即并入
        self._set_record()
        f1, f2 = self._patch_engine({"pid": 4242, "source": "pbf"})
        with f1, f2:
            self.assertEqual(wa._excel_owner_pid(None), 4242)

    def test_missing_record_file_ignored(self):
        started = time.time() - 5
        f1, f2 = self._patch_engine({"pid": 4242, "source": "pbf"})
        with f1, f2:                                  # owner_file 指向不存在的文件
            self.assertIsNone(wa._excel_owner_pid(started))

    def test_engine_unavailable_returns_none(self):
        # engine 不可导入（轻量环境）：退化为不并入，只按基线判断
        with mock.patch.dict(sys.modules, {"engine.bootstrap": None}):
            self.assertIsNone(wa._excel_owner_pid(None))


if __name__ == "__main__":
    unittest.main()
