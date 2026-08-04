from collections.abc import Callable

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import virtual_finger
from .virtual_finger import as_motion


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)


MAGICAL_SEQUENCE = [as_motion(elem) for elem in ("LWin-S-M", "U-Alt")]


def run(
    func: Callable,
    finished: Callable | None = None,
    focus_changed_in_subthread: bool = False,
) -> None:
    if focus_changed_in_subthread:
        finger = virtual_finger.VirtualFinger(0)
        finger.send_motion_sequence(*MAGICAL_SEQUENCE)

    def _finished(job_item: ckit.JobItem) -> None:
        keymap.setInput_Modifier(0)
        if finished is not None:
            finished(job_item)

    job = ckit.JobItem(func, _finished)
    ckit.JobQueue.defaultQueue().enqueue(job)
