from collections.abc import Callable

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import subthread, virtual_finger
from .common import CallbackFunc, balloon, delay
from .virtual_finger import Tap


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)
    subthread.setup(keymap)


def get_string() -> str:
    try:
        return ckit.getClipboardText() or ""
    except Exception:  # noqa: BLE001
        return ""


def get_latest_clipboard_history() -> str:
    try:
        return keymap.clipboard_history.items[0]
    except IndexError:
        return ""


def set_string(s: str) -> None:
    try:
        ckit.setClipboardText(str(s))
    except Exception:  # noqa: BLE001, S110
        pass


class Manager:
    tap_to_register: Tap
    tap_to_paste: Tap = Tap("C-V")
    terminal_process: tuple[str, ...] = (
        "pwsh.exe",
        "powershell.exe",
        "wezterm-gui.exe",
    )
    finger: virtual_finger.VirtualFinger

    def __init__(self, cut_mode: bool = False) -> None:
        if cut_mode:
            self.tap_to_register = Tap("C-X")
        else:
            self.tap_to_register = Tap("C-C")
        self.finger = virtual_finger.VirtualFinger()

    def send_register_key(self) -> None:
        self.finger.send_compiled(self.tap_to_register)

    def send_paste_key(self) -> None:
        self.finger.send_compiled(self.tap_to_paste)

    def paste(
        self,
        s: str | None = None,
        format_func: Callable[[str], str] | None = None,
    ) -> None:
        if s is None:
            s = get_string()
            if any(0x10000 < ord(c) for c in s):
                # newer emoji
                self.send_paste_key()
                return

            if len(s) < 1:
                # empty clipboard could be image.
                self.send_paste_key()
                return
        if format_func is not None:
            s = format_func(s)
        if keymap.getWindow().getProcessName() in self.terminal_process:
            s = s.strip()
        set_string(s)
        self.send_paste_key()

    def after_register(self, deferred: Callable[[ckit.JobItem], None]) -> None:
        cb = get_latest_clipboard_history()
        self.send_register_key()
        delay(40)

        def _watch_clipboard(job_item: ckit.JobItem) -> None:
            job_item.origin = cb
            job_item.copied = ""
            trial = 600
            for _ in range(trial):
                s = get_latest_clipboard_history()
                if not s.strip():
                    continue
                if s != job_item.origin:
                    job_item.copied = s
                    break

        subthread.run(_watch_clipboard, deferred)


class FIFOStack:
    def __init__(self) -> None:
        self.items = []
        self.enabled = False

    def _enable(self) -> None:
        balloon(keymap, "FIFO mode ON!")
        self.enabled = True

    def _disable(self, alert: bool = True) -> None:
        if alert:
            balloon(keymap, "FIFO mode OFF!")
        self.enabled = False

    def toggle(self) -> None:
        if self.enabled:
            self._disable()
        else:
            self._enable()

    def register(self, s: str) -> None:
        if self.enabled:
            self.items.append(s)
            msg = f"FIFO stack total: {self.count}"
            balloon(keymap, msg)
        else:
            balloon(keymap, "FIFO mode is not enabled.")

    def reset(self) -> None:
        self.items = []

    def bulk_register(self, lines: str) -> None:
        if self.enabled:
            self.reset()
            self.items = [line for line in lines.splitlines() if line.strip()]
            msg = f"FIFO stack total: {self.count}"
            balloon(keymap, msg)
        else:
            balloon(keymap, "FIFO mode is not enabled.")

    def bulk_paste(self, delimiter: str) -> str:
        if self.enabled:
            s = delimiter.join(self.items)
            self.reset()
            self._disable()
            return s
        return ""

    def join_items(self, sep: str) -> str:
        if not self.enabled:
            balloon(keymap, "FIFO mode is not enabled.")
            return ""
        s = sep.join(self.items)
        self.reset()
        self._disable()
        return s

    @property
    def count(self) -> int:
        return len(self.items)

    def pop(self) -> str | None:
        if not self.enabled:
            balloon(keymap, "FIFO mode is not enabled.")
            return None
        if 0 < self.count:
            cb = self.items.pop(0)
            if self.count == 0:
                balloon(keymap, "FIFO mode OFF! (pasted last item)", 5000)
                self._disable(False)
            else:
                balloon(keymap, f"FIFO next:{self.items[0]}", 5000)
            return cb
        return None


def remove_whitespace(s: str) -> str:
    return s.strip().translate(
        str.maketrans(
            "",
            "",
            "\u0009\u0020\u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c\u200d\u200e\u200f\u202f\u205f\u3000\ufeff",
        )
    )


def invoke_clean_paster(no_space: bool = False, no_break: bool = False) -> CallbackFunc:
    def _clean(s) -> str:
        s = s.strip()
        if no_space:
            s = remove_whitespace(s)
        if no_break:
            s = "".join(s.splitlines())
        return s

    def _paste() -> None:
        Manager().paste(format_func=_clean)

    return _paste


def simple_quote(s: str) -> str:
    lines = s.strip().splitlines()
    return "\n".join([keymap.quote_mark + line for line in lines])


def as_single_line(s: str) -> str:
    lines = s.strip().splitlines()
    return keymap.quote_mark + "".join([line.strip() for line in lines])


def skip_blank_line(s: str) -> str:
    lines = []
    for line in s.strip().splitlines():
        if 0 < len(line.strip()):
            lines.append(keymap.quote_mark + line)
        else:
            lines.append("")
    return "\n".join(lines)


def invoke_quote_paster(func: Callable[[str], str]) -> CallbackFunc:
    def _paster() -> None:
        Manager().paste(None, func)

    return _paster
