from typing import Callable

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore  # noqa: F403

from . import virtual_finger

keymap: WindowKeymap = None

virtual_finger.keymap = keymap

MAGICAL_KEY = virtual_finger.VirtualFinger().compile("LWin-S-M", "U-Alt")


def run(
    func: Callable,
    finished: Callable | None = None,
    focus_changed_in_subthread: bool = False,
) -> None:

    finger = virtual_finger.VirtualFinger(0)
    if focus_changed_in_subthread:
        finger.send_compiled(*MAGICAL_KEY)

    def _finished(job_item: ckit.JobItem) -> None:
        keymap.setInput_Modifier(0)
        if finished is not None:
            finished(job_item)

    job = ckit.JobItem(func, _finished)
    ckit.JobQueue.defaultQueue().enqueue(job)
