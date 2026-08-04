import subprocess
import webbrowser

import ckit  # type: ignore
import pyauto  # type: ignore

from keyhac import *  # type: ignore

from .tools import subthread as subthread_tool
from .tools import virtual_finger as vf_tool
from .tools import window_activate as window_activate_tool
from .tools.browser_info import SystemBrowser
from .tools.common import (
    balloon,
    check_fzf,
    delay,
    is_global_target,
    is_keyhac_console,
    open_vscode,
)
from .tools.window_activate import WindowActivator, WndScanner


def _open_vscode() -> None:
    open_vscode()


def fuzzy_window_switcher(keymap) -> None:
    if not check_fzf():
        balloon(keymap, "cannot find fzf on PC.")
        return

    ignore_list = [
        "fzf.exe",
        "explorer.exe",
        "MouseGestureL.exe",
        "TextInputHost.exe",
        "SystemSettings.exe",
        "ApplicationFrameHost.exe",
    ]

    def _fzf_wnd(job_item: ckit.JobItem) -> None:
        job_item.result = None
        delay()
        popup_table = {}

        proc = subprocess.Popen(
            ["fzf.exe", "--no-mouse", "--margin=1"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            creationflags=subprocess.HIGH_PRIORITY_CLASS,
        )

        def _walk(wnd: pyauto.Window, _) -> bool:
            if not wnd:
                return False
            if not wnd.isVisible():
                return True
            if not wnd.isEnabled():
                return True
            if is_keyhac_console(wnd):
                return True
            if wnd.getProcessName() in ignore_list:
                return True
            if not wnd.getText():
                return True
            if popup := wnd.getLastActivePopup():
                n = popup.getProcessName().replace(".exe", "")
                if t := popup.getText():
                    n += f"[{t}]"
                popup_table[n] = popup
                if proc.stdin:
                    proc.stdin.write(n + "\n")
            return True

        try:
            pyauto.Window.enum(_walk, None)
            if proc.stdin:
                proc.stdin.close()
        except Exception as e:  # noqa: BLE001
            print(e)
            return

        result, err = proc.communicate()
        if proc.returncode != 0:
            if err:
                print(err)
            return
        result = result.strip()
        if len(result) < 1:
            return
        wnd = popup_table.get(result, None)
        if wnd is not None:
            job_item.result = WindowActivator(wnd).activate()

    def _finished(job_item: ckit.JobItem) -> None:
        if job_item.result is None:
            return
        if not job_item.result:
            vf_tool.VirtualFinger().send("LCtrl-LAlt-Tab")

    subthread_tool.run(_fzf_wnd, _finished, True)


SYSTEM_BROWSER = SystemBrowser()


def search_on_browser(keymap) -> None:
    finger = vf_tool.VirtualFinger(20)
    if keymap.getWindow().getProcessName() == SYSTEM_BROWSER.get_exe_name():
        finger.send("C-T")
        return

    def _activate(job_item: ckit.JobItem) -> None:
        delay()
        job_item.result = None
        scanner = WndScanner(
            SYSTEM_BROWSER.get_exe_name(),
            SYSTEM_BROWSER.get_wnd_class(),
        )
        scanner.scan()
        wnd = scanner.found
        if wnd is None:
            webbrowser.open("http://")
        else:
            job_item.result = WindowActivator(wnd).activate()

    def _finished(job_item: ckit.JobItem) -> None:
        if job_item.result is not None:
            if job_item.result:
                finger.send("C-T")
            else:
                finger.send("LCtrl-LAlt-Tab")

    subthread_tool.run(_activate, _finished, True)


def bind(keymap) -> None:

    vf_tool.setup(keymap)
    subthread_tool.setup(keymap)
    window_activate_tool.setup(keymap)

    # keymap working on any window
    km = keymap.defineWindowKeymap(check_func=is_global_target)

    REMAP_SINGLE_KEY = {
        "U1-F": (
            "cfiler.exe",
            "CfilerWindowClass",
            r"${USERPROFILE}\Personal\portable_apps\cfiler\cfiler.exe",
        ),
        "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
        "U1-K": ("KIRI10.exe", "*"),
        "C-U1-S": ("smoothcsv-app.exe", "*"),
        "LC-U1-M": (
            "Mery.exe",
            "TChildForm",
            r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
        ),
        "LC-U1-N": (
            "notepad.exe",
            "Notepad",
            r"C:\Windows\System32\notepad.exe",
        ),
        "LC-AtMark": (
            "wezterm-gui.exe",
            "org.wezfurlong.wezterm",
            r"${USERPROFILE}\scoop\apps\wezterm\current\wezterm-gui.exe",
        ),
    }

    for key, params in REMAP_SINGLE_KEY.items():
        func = window_activate_tool.invoke_cute_exec(*params)
        km[key] = func

    REMAP_KEY_SEQUENCE = {
        "Space": (
            SYSTEM_BROWSER.get_exe_name(),
            SYSTEM_BROWSER.get_wnd_class(),
            SYSTEM_BROWSER.get_exe_path(),
        ),
        "C": (
            "chrome.exe",
            "Chrome_WidgetWin_1",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        ),
        "D": (
            "vivaldi.exe",
            "Chrome_WidgetWin_1",
            r"${LOCALAPPDATA}\Vivaldi\Application\vivaldi.exe",
        ),
        "S": (
            "slack.exe",
            "Chrome_WidgetWin_1",
            r"${LOCALAPPDATA}\slack\slack.exe",
        ),
        "F": (
            "firefox.exe",
            "MozillaWindowClass",
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
        ),
        "B": (
            "thunderbird.exe",
            "MozillaWindowClass",
            r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
        ),
        "K": (
            "ksnip.exe",
            "Qt5152QWindowIcon",
            r"${USERPROFILE}\scoop\apps\ksnip\current\ksnip.exe",
        ),
        "O": ("Obsidian.exe", "Chrome_WidgetWin_1"),
        "P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
        "C-P": ("powerpnt.exe", "PPTFrameClass"),
        "E": ("EXCEL.EXE", "XLMAIN"),
        "W": ("WINWORD.EXE", "OpusApp"),
        "V": ("Code.exe", "Chrome_WidgetWin_1", _open_vscode),
        "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1"),
        "M": (
            "Mery.exe",
            "TChildForm",
            r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
        ),
        "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
    }

    km["U1-C"] = keymap.defineMultiStrokeKeymap()

    for key, params in REMAP_KEY_SEQUENCE.items():
        func = window_activate_tool.invoke_cute_exec(*params)
        km["U1-C"][key] = func

    km["U1-E"] = lambda: fuzzy_window_switcher(keymap)

    km["U0-Q"] = lambda: search_on_browser(keymap)
