"""AO 登录辅助 — 检测 SAP Analysis 登录框，自动填账号密码并处理安全警告。

现场实测的元素定位（TCL BPC PRD 登录框，WinForms）：
- 窗口标题：登录到系统 TCL BPC PRD
- 客户端 Edit  AutomationId=1969556
- 用户   Edit  AutomationId=1510560
- 密码   Edit  AutomationId=2821500
- 语言   Edit  AutomationId=2230944
- 确定   Button AutomationId=1641710
- 连接不安全时会二次弹窗，标题"不安全连接"，按钮文本"继续"

注意：该登录框不支持 ValuePattern / InvokePattern，只能先点击再 SendKeys。
"""
from __future__ import annotations

import logging
import time

import uiautomation as auto

from .sendkeys import escape_sendkeys

log = logging.getLogger(__name__)

LOGIN_TITLE = "登录到系统 TCL BPC PRD"
INSECURE_TITLE = "不安全连接"

EDIT_CLIENT = "_clientTextBox"
EDIT_USER = "_userTextBox"
EDIT_PASS = "_passwordTextBox"
EDIT_LANG = "_languageTextBox"
BTN_OK = "_okButton"
BTN_CANCEL = "_cancelButton"


class AoLogin:
    def __init__(self, username: str, password: str, client: str = "333", lang: str = "zh") -> None:
        self.username = username
        self.password = password
        self.client = client
        self.lang = lang
        # ensure() 期间是否见过登录框：区分"密码错误/登录框卡住"与"从未弹框"
        self.saw_login_box = False

    # ---------- 元素查找 ----------

    @staticmethod
    def _find_login(timeout: float = 0.5):
        try:
            return auto.WindowControl(Name=LOGIN_TITLE, searchDepth=2).Exists(timeout)
        except Exception:                 # UIA 偶发 COMError（窗口正在销毁等）
            return False

    @staticmethod
    def _find_insecure(timeout: float = 0.5):
        try:
            return auto.WindowControl(Name=INSECURE_TITLE, searchDepth=2).Exists(timeout)
        except Exception:
            return False

    @staticmethod
    def _click(ctrl) -> None:
        """点击控件。WinForms 的 Pane 不支持 Click()，退回坐标点击。"""
        try:
            ctrl.Click(simulateMove=False)
            return
        except Exception:
            pass
        r = ctrl.BoundingRectangle
        x = (r.left + r.right) // 2
        y = (r.top + r.bottom) // 2
        auto.SetCursorPos(x, y)
        time.sleep(0.15)
        auto.Click()

    @classmethod
    def _type_into(cls, ctrl, text: str) -> None:
        cls._click(ctrl)
        time.sleep(0.3)
        auto.SendKeys("{Ctrl}a")
        time.sleep(0.2)
        auto.SendKeys("{Delete}")
        time.sleep(0.1)
        if text:
            # 密码可能含 + ^ % ~ ( ) { } —— 不转义会被 SendKeys 当控制字符
            auto.SendKeys(escape_sendkeys(text))
        time.sleep(0.3)

    @staticmethod
    def _by_id(dlg, auto_id: str):
        """按 AutomationId 查找输入框。

        uiautomation 读到的 AutomationId 是语义名（如 _passwordTextBox），
        与 PowerShell UIAutomation 显示的纯数字 ID 不同。
        """
        return dlg.EditControl(AutomationId=auto_id, searchDepth=10)

    @staticmethod
    def _btn_by_id(dlg, auto_id: str):
        return dlg.ButtonControl(AutomationId=auto_id, searchDepth=10)

    # ---------- 具体动作 ----------

    def _fill_and_submit(self) -> bool:
        try:
            return self._fill_and_submit_inner()
        except Exception as exc:
            log.warning(f"填写登录框时出错（稍后重试）: {exc}")
            return False

    def _fill_and_submit_inner(self) -> bool:
        dlg = auto.WindowControl(Name=LOGIN_TITLE, searchDepth=2)
        if not dlg.Exists(1):
            return False

        filled = 0
        for auto_id, value, label in (
            (EDIT_CLIENT, self.client, "客户端"),
            (EDIT_USER, self.username, "用户"),
            (EDIT_LANG, self.lang, "语言"),
            (EDIT_PASS, self.password, "密码"),
        ):
            ctrl = self._by_id(dlg, auto_id)
            if not ctrl.Exists(1):
                log.warning(f"登录框缺少[{label}]输入框")
                continue
            self._type_into(ctrl, value)
            filled += 1
            log.info(f"已填写[{label}]")

        ok = self._btn_by_id(dlg, BTN_OK)
        if not ok.Exists(1):
            log.error(f"登录框未找到[确定]按钮（已填 {filled}/4 项）")
            return False
        self._click(ok)
        log.info(f"已点击登录[确定]（已填 {filled}/4 项）")
        return True

    def _accept_insecure(self) -> bool:
        """处理"不安全连接"警告。

        实测：该对话框标题栏为空，按窗口标题找不到，只能全局搜索"继续"按钮。
        UIA 在窗口刚消失时可能抛 COMError，这里统一当"没有该弹窗"处理。
        """
        try:
            btn = auto.ButtonControl(Name="继续", searchDepth=8)
            if not btn.Exists(0.8):
                return False
            self._click(btn)
            log.info("已确认[不安全连接]警告")
            return True
        except Exception as exc:
            log.debug(f"检测[不安全连接]时跳过: {exc}")
            return False

    # ---------- 主流程 ----------

    def ensure(self, timeout: float = 60.0, expect_login: bool = False) -> bool:
        """确保已登录。返回 True=已登录/无需登录，False=超时仍停留在登录框。

        - expect_login=False（默认）：若整个轮询期间从未出现登录框，判定已用缓存凭据
          登录并尽快返回 True（避免已登录时白等）。
        - expect_login=True：登录框必然会出现（打开导航表后 AO 要加载 ~10s 才弹框），
          必须等到它出现并填掉，不能提前判定已登录。
        """
        deadline = time.time() + timeout
        saw_login = False

        while time.time() < deadline:
            try:
                # 任何时刻出现"不安全连接"都先处理掉
                if self._accept_insecure():
                    time.sleep(3)
                    continue

                if self._find_login(0.5):
                    saw_login = True
                    self.saw_login_box = True
                    log.info("检测到 AO 登录框，开始自动登录")
                    self._fill_and_submit()
                    time.sleep(5)

                    # 点确定后通常立刻弹"不安全连接"
                    if self._accept_insecure():
                        time.sleep(5)

                    if not self._find_login(0.5):
                        log.info("登录成功")
                        return True
                    log.warning("登录框仍在，重试")
                    continue

                # 没检测到登录框
                if saw_login:
                    # 登录框曾经出现但现在消失 → 登录成功
                    log.info("登录框已消失，登录成功")
                    return True

                # 从未出现过登录框：仅当不要求登录框出现时，才允许提前判定已登录
                if not expect_login:
                    time.sleep(2)
                    if not self._find_login(0.5):
                        log.info("未出现登录框，判定已登录")
                        return True
            except Exception as exc:      # UIA 偶发 COMError，忽略后继续轮询
                log.debug(f"登录轮询异常（忽略并重试）: {exc}")

            time.sleep(1)

        log.error(f"等待登录超时（{timeout}s）")
        return False
