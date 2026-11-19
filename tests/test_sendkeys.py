"""SendKeys 特殊字符转义。"""
import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _load(name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / "engine" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


sendkeys = _load("sendkeys")


class EscapeSendKeysTest(unittest.TestCase):
    def test_plain_text_unchanged(self):
        self.assertEqual(sendkeys.escape_sendkeys("abc123@._-"), "abc123@._-")

    def test_ctrl_and_alt_chars_escaped(self):
        self.assertEqual(
            sendkeys.escape_sendkeys("a+b^c%d~e"),
            "a{+}b{^}c{%}d{~}e")

    def test_parentheses_escaped(self):
        self.assertEqual(sendkeys.escape_sendkeys("p(w)d"), "p{(}w{)}d")

    def test_braces_escaped(self):
        self.assertEqual(sendkeys.escape_sendkeys("a{b}c"), "a{{}b{}}c")

    def test_password_example(self):
        self.assertEqual(
            sendkeys.escape_sendkeys("P@ss{1}+Ab^"),
            "P@ss{{}1{}}{+}Ab{^}")


if __name__ == "__main__":
    unittest.main()
