import fnmatch

import ckit  # type: ignore
import pyauto  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import subthread, virtual_finger
from .common import CallbackFunc, delay, shell_exec


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)
    subthread.setup(keymap)


class WndScanner:
    def __init__(self, exe_name: str, class_name: str = "") -> None:
        self.exe_name = exe_name
        self.class_name = class_name
        self.found = None

    def reset(self) -> None:
        self.found = None

    def scan(self) -> None:
        self.reset()
        pyauto.Window.enum(self.walk, None)

    def walk(self, wnd: pyauto.Window, _) -> bool:
        if not wnd:
            return False
        if not wnd.isVisible():
            return True
        if not wnd.isEnabled():
            return True
        if self.class_name and not fnmatch.fnmatch(wnd.getClassName(), self.class_name):
            return True
        if not fnmatch.fnmatch(wnd.getProcessName(), self.exe_name):
            return True
        popup = wnd.getLastActivePopup()
        if not popup:
            return True
        self.found = popup
        return False


class WindowActivator:
    def __init__(self, wnd: pyauto.Window) -> None:
        self._target = wnd

    def _check(self) -> bool:
        return pyauto.Window.getForeground() == self._target

    def activate(self) -> bool:
        if self._check():
            return True

        if self._target.isMinimized():
            self._target.restore()
            delay()

        interval = 20
        trial = 40
        for _ in range(trial):
            try:
                self._target.setForeground()
                delay(interval)
                if self._check():
                    self._target.setForeground(True)
                    return True
            except Exception as e:  # noqa: BLE001
                print("Failed to activate window due to exception:", e)
                return False

        print("Failed to activate window due to timeout.")
        return False


def invoke_cute_exec(
    exe_name: str, class_name: str = "", exe_path: str | CallbackFunc = ""
) -> CallbackFunc:
    def _executor() -> None:
        scanner = WndScanner(exe_name, class_name)

        def _activate(job_item: ckit.JobItem) -> None:
            job_item.result = None
            delay(40)
            scanner.scan()
            wnd = scanner.found
            if wnd is None:
                if exe_path:
                    if isinstance(exe_path, str):
                        shell_exec(exe_path)
                    else:
                        exe_path()
            else:
                job_item.result = WindowActivator(wnd).activate()

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.result is None:
                return
            if not job_item.result:
                virtual_finger.VirtualFinger().send("LCtrl-LAlt-Tab")

        subthread.run(_activate, _finished, True)

    return _executor
