import os
import shutil
from pathlib import Path
from typing import NamedTuple

from keyhac import ActivateApplication  # ty: ignore[unresolved-import]
from libs.system_browser import get_browser_path


class LaunchEntry(NamedTuple):
    app: str
    launch: str | None

    def as_params(self) -> dict:
        params = {"app": self.app}
        if self.launch is not None:
            params["launch"] = os.path.expandvars(self.launch)
        return params


SINGLE_KEY_MAPPING: dict[str, LaunchEntry] = {
    "U1-F": LaunchEntry(
        "cfiler.exe",
        r"${USERPROFILE}\Personal\portable_apps\cfiler\cfiler.exe",
    ),
    "U1-P": LaunchEntry("SumatraPDF.exe", None),
    "C-U1-S": LaunchEntry("smoothcsv-app.exe", None),
    "LC-U1-M": LaunchEntry(
        "Mery.exe",
        r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
    ),
    "LC-U1-N": LaunchEntry(
        "notepad.exe",
        r"C:\Windows\System32\notepad.exe",
    ),
    "LC-AtMark": LaunchEntry(
        "wezterm-gui.exe",
        r"${USERPROFILE}\scoop\apps\wezterm\current\wezterm-gui.exe",
    ),
}

MULTI_KEY_MAPPING: dict[str, LaunchEntry] = {
    "C": LaunchEntry(
        "chrome.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    ),
    "S": LaunchEntry(
        "slack.exe",
        r"${LOCALAPPDATA}\slack\slack.exe",
    ),
    "F": LaunchEntry(
        "firefox.exe",
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
    ),
    "B": LaunchEntry(
        "thunderbird.exe",
        r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
    ),
    "P": LaunchEntry("SumatraPDF.exe", None),
    "E": LaunchEntry("EXCEL.EXE", None),
    "W": LaunchEntry("WINWORD.EXE", None),
    "M": LaunchEntry(
        "Mery.exe",
        r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
    ),
    "X": LaunchEntry("explorer.exe", None),
}

if (code_path := shutil.which("code")) is not None:
    MULTI_KEY_MAPPING["V"] = LaunchEntry("Code.exe", code_path)

if (browser_path := get_browser_path()) is not None:
    MULTI_KEY_MAPPING["Space"] = LaunchEntry(Path(browser_path).name, browser_path)


def bind(keymap, kt) -> None:

    for key, entry in SINGLE_KEY_MAPPING.items():
        kt[key] = ActivateApplication(**entry.as_params())

    kt_second_key = keymap.define_keytable()
    for key, entry in MULTI_KEY_MAPPING.items():
        kt_second_key[key] = ActivateApplication(**entry.as_params())

    kt["U1-C"] = kt_second_key
