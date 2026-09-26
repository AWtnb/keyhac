from collections.abc import Callable

from keyhac import ThreadedAction  # ty: ignore[unresolved-import]

from . import sender
from ._common import delay


def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap
    sender.setup(keymap)


def get_string() -> str:
    return keymap.clipboard.get_text() or ""


def get_latest_clipboard_history() -> str:
    return keymap.clipboard_history.get_current() or ""


def set_string(s: str) -> None:
    keymap.clipboard.set_text(s)


def send_copy_key() -> None:
    sender.send_keys("C-C")


def send_paste_key() -> None:
    sender.send_keys("C-V")


class Paste(ThreadedAction):
    def __init__(
        self, text: str | None = None, format_func: Callable[[str], str] | None = None
    ) -> None:
        self.text = text
        self.format_func = format_func

    def run(self) -> None:
        if self.text is None:
            s = get_string()
            if s:
                if self.format_func is not None:
                    s = self.format_func(s)
                set_string(s)
        else:
            s = self.text
            if self.format_func is not None:
                s = self.format_func(s)
            set_string(s)
        delay()

    def finished(self, _):
        send_paste_key()


class CopyThen(ThreadedAction):
    def __init__(self, deferred: Callable[[str], None]):
        self.deferred = deferred
        self.origin = ""

    def starting(self):
        self.origin = get_latest_clipboard_history()
        send_copy_key()

    def run(self) -> str:
        delay(40)
        trial = 40
        for _ in range(trial):
            s = get_latest_clipboard_history()
            if not s.strip():
                continue
            if s != self.origin:
                return s
        return self.origin

    def finished(self, result):
        self.deferred(result)
