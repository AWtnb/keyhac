from collections.abc import Callable

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # ty: ignore[unresolved-import]

from . import subthread, virtual_finger
from .common import CallbackFunc, delay
from .str_tools import remove_whitespace
from .virtual_finger import Tap


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)
    subthread.setup(keymap)

    global VF  # ty: ignore[unresolved-global]
    VF = virtual_finger.VirtualFinger()


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


TAP_TO_COPY = Tap("C-C")
TAP_TO_PASTE = Tap("C-V")


def send_copy_key() -> None:
    VF.send_compiled(TAP_TO_COPY)


def send_paste_key() -> None:
    VF.send_compiled(TAP_TO_PASTE)


def paste(
    s: str | None = None, format_func: Callable[[str], str] | None = None
) -> None:
    if s is None:
        s = get_string()
        if any(0x10000 < ord(c) for c in s):
            # newer emoji
            send_paste_key()
            return

        if len(s) < 1:
            # empty clipboard could be image.
            send_paste_key()
            return

    if format_func is not None:
        s = format_func(s)

    set_string(s)
    send_paste_key()


def copy_then(deferred: Callable[[ckit.JobItem], None]) -> None:
    cb = get_latest_clipboard_history()
    virtual_finger.VirtualFinger().send_compiled(Tap("C-C"))
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
        paste(format_func=_clean)

    return _paste
