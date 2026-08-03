from collections.abc import Callable

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # ty: ignore[unresolved-import]

from . import subthread, virtual_finger
from .common import CallbackFunc, delay
from .format_str import remove_whitespace
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
