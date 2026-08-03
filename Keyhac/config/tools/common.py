import datetime
import os
import shutil
import subprocess
import time
import webbrowser
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pyauto  # type: ignore

from keyhac import *  # type: ignore


def get_now() -> datetime.datetime:
    JST = datetime.timezone(datetime.timedelta(hours=9))
    return datetime.datetime.now(tz=JST)


def balloon(keymap, message: str | Exception, timeout_msec: int = 1500) -> None:
    title = get_now().strftime("%Y%m%d-%H%M%S-%f")
    print(message)
    try:
        keymap.popBalloon(title, message, timeout_msec)
    except Exception as e:  # noqa: BLE001
        print(e)


def smart_check_path(path: str | Path, timeout_sec: float | None = None) -> bool:
    """
    CASE-INSENSITIVE path check with timeout
    """
    p = path if isinstance(path, Path) else Path(path)
    try:
        future = ThreadPoolExecutor(max_workers=1).submit(p.exists)
        return future.result(timeout_sec)
    except Exception:  # noqa: BLE001
        return False


def check_fzf() -> bool:
    return shutil.which("fzf.exe") is not None


def open_vscode(*args: str) -> bool:
    try:
        if code_path := shutil.which("code"):
            cmd = [code_path] + list(args)
            subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=False)
            return True
        return False
    except Exception as e:  # noqa: BLE001
        print(e)
        return False


def is_file_locked(path: Path | str) -> bool:
    try:
        with open(path, "a"):
            return False
    except OSError:
        return True


def resolve_scoop_shim(path: str) -> str:
    if r"scoop\shims" in path and path.lower().endswith(".exe"):
        real = str(
            Path(path)
            .with_suffix(".shim")
            .read_text()
            .strip()
            .split(" = ")[-1]
            .replace('"', "")
        )
        return real
    return path


def shell_exec(path: str, *args) -> None:
    if not isinstance(path, str):
        path = str(path)
    if path.startswith("http"):
        webbrowser.open(path)
        return
    path = os.path.expandvars(path)
    try:
        cmd = ["start", "", path] + list(args)
        subprocess.run(cmd, shell=True, check=False)
    except Exception as e:  # noqa: BLE001
        print(e)


CallbackFunc = Callable[[], None]


def delay(msec: int = 50) -> None:
    if 0 < msec:
        time.sleep(msec / 1000)


def is_browser(wnd: pyauto.Window) -> bool:
    return wnd.getProcessName() in ("chrome.exe", "vivaldi.exe", "firefox.exe")


def is_global_target(wnd: pyauto.Window) -> bool:
    return not (is_browser(wnd) and wnd.getText().startswith("ESET - "))


def is_keyhac_console(wnd: pyauto.Window) -> bool:
    return wnd.getProcessName() == "keyhac.exe" and not wnd.getFirstChild()
