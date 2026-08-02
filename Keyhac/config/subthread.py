from typing import Callable

import ckit  # type: ignore

from .virtual_finger import VirtualFinger


def run(
    keymap,
    func: Callable,
    finished: Callable | None = None,
    focus_changed_in_subthread: bool = False,
) -> None:
    MAGICAL_KEY = VirtualFinger(keymap).compile("LWin-S-M", "U-Alt")

    finger = VirtualFinger(keymap, 0)
    if focus_changed_in_subthread:
        finger.send_compiled(*MAGICAL_KEY)

    def _finished(job_item: ckit.JobItem) -> None:
        keymap.setInput_Modifier(0)
        if finished is not None:
            finished(job_item)

    job = ckit.JobItem(func, _finished)
    ckit.JobQueue.defaultQueue().enqueue(job)
