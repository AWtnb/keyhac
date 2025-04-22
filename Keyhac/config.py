import datetime
import os
import fnmatch
import re
import time
import shutil
import subprocess
import urllib.parse
import unicodedata
from typing import Union, Callable, Dict, List, NamedTuple
from pathlib import Path
from winreg import HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, OpenKey, QueryValueEx
from concurrent.futures import ThreadPoolExecutor

import ckit
import pyauto
from keyhac import *
from keyhac_keymap import Keymap, KeyCondition, WindowKeymap
from keyhac_listwindow import ListWindow


def smart_check_path(path: Union[str, Path], timeout_sec: Union[int, float, None] = None) -> bool:
    """CASE-INSENSITIVE path check with timeout"""
    p = path if type(path) is Path else Path(path)
    try:
        future = ThreadPoolExecutor(max_workers=1).submit(p.exists)
        return future.result(timeout_sec)
    except:
        return False


def check_fzf() -> bool:
    return shutil.which("fzf.exe") is not None


def configure(keymap):

    def balloon(message: str) -> None:
        title = datetime.datetime.today().strftime("%Y%m%d-%H%M%S-%f")
        print(message)
        try:
            keymap.popBalloon(title, message, 1500)
        except:
            pass

    def shell_exec(path: str, *args) -> None:
        if type(path) is not str:
            path = str(path)
        params = []
        for arg in args:
            if len(arg.strip()):
                if " " in arg:
                    params.append('"{}"'.format(arg))
                else:
                    params.append(arg)
        try:
            pyauto.shellExecute(None, path, " ".join(params), "")
        except:
            balloon("invalid path: '{}'".format(path))

    def subthread_run(func: Callable, finished: Union[Callable, None] = None) -> None:
        job = ckit.JobItem(func, finished)
        ckit.JobQueue.defaultQueue().enqueue(job)

    ################################
    # general setting
    ################################

    class PathHandler:
        def __init__(self, path: str) -> None:
            self._path = os.path.expandvars(path)

        @property
        def path(self) -> str:
            return self._path

        def is_accessible(self) -> bool:
            return smart_check_path(self._path)

        @staticmethod
        def args_to_param(args: tuple) -> str:
            params = []
            for arg in args:
                if 0 < len(arg.strip()):
                    if " " in arg:
                        params.append('"{}"'.format(arg))
                    else:
                        params.append(arg)
            return " ".join(params)

        def run(self, *args) -> None:
            if self.is_accessible():
                shell_exec(self._path, *args)
            else:
                balloon("invalid-path: '{}'".format(self._path))

    def get_editor() -> str:
        vscode_path = PathHandler(r"${USERPROFILE}\scoop\apps\vscode\current\Code.exe")
        if vscode_path.is_accessible():
            return vscode_path.path
        return "notepad.exe"

    KEYHAC_EDITOR = get_editor()

    # console theme
    keymap.setFont("HackGen", 16)

    def set_custom_theme():
        name = "black"

        custom_theme = {
            "bg": (63, 59, 57),
            "fg": (160, 180, 167),
            "cursor0": (255, 255, 255),
            "cursor1": (255, 64, 64),
            "bar_fg": (0, 0, 0),
            "bar_error_fg": (255, 64, 64),
            "select_bg": (223, 244, 119),
            "select_fg": (63, 59, 57),
            "caret0": (255, 255, 255),
            "caret1": (255, 0, 0),
        }
        ckit.ckit_theme.theme_name = name

        for k, v in custom_theme.items():
            ckit.ckit_theme.ini.set("COLOR", k, str(v))
        keymap.console_window.reloadTheme()

    set_custom_theme()

    # user modifier
    keymap.replaceKey("(29)", 235)  # "muhenkan" => 235
    keymap.replaceKey("(28)", 236)  # "henkan" => 236
    keymap.defineModifier(235, "User0")  # "muhenkan" => "U0"
    keymap.defineModifier(236, "User1")  # "henkan" => "U1"

    # enable clipbard history
    keymap.clipboard_history.enableHook(True)

    # history max size
    keymap.clipboard_history.maxnum = 200
    keymap.clipboard_history.quota = 10 * 1024 * 1024

    # quote mark when paste with Ctrl.
    keymap.quote_mark = "> "

    ################################
    # key remap
    ################################

    class CheckWnd:
        @staticmethod
        def is_browser(wnd: pyauto.Window) -> bool:
            return wnd.getProcessName() in ("chrome.exe", "vivaldi.exe", "firefox.exe")

        @classmethod
        def is_global_target(cls, wnd: pyauto.Window) -> bool:
            return not (cls.is_browser(wnd) and wnd.getText().startswith("ESET - "))

        @staticmethod
        def is_keyhac_console(wnd: pyauto.Window) -> bool:
            return wnd.getProcessName() == "keyhac.exe" and not wnd.getFirstChild()

    # keymap working on any window
    keymap_global = keymap.defineWindowKeymap(check_func=CheckWnd().is_global_target)

    # clipboard menu
    keymap_global["LC-LS-X"] = keymap.command_ClipboardList

    # keyboard macro
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["S-U0-0"] = keymap.command_RecordClear
    keymap_global["U1-0"] = keymap.command_RecordPlay
    keymap_global["U1-F4"] = keymap.command_RecordPlay
    keymap_global["C-U0-0"] = keymap.command_RecordPlay

    class KeyAllocator:
        mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
        key_status = ("D-", "U-")

        def __init__(self, km: WindowKeymap) -> None:
            self._keymap = km

        def cursor_keys(self) -> None:
            for mod_key in self.mod_keys:
                for key, value in {
                    # move cursor
                    "H": "Left",
                    "J": "Down",
                    "K": "Up",
                    "L": "Right",
                    # Back / Delete
                    "B": "Back",
                    "D": "Delete",
                    # Home / End
                    "A": "Home",
                    "E": "End",
                    # Enter
                    "Space": "Enter",
                }.items():
                    self._keymap[mod_key + "U0-" + key] = mod_key + value

        def apply(self, mapping_dict: dict) -> None:
            for key, value in mapping_dict.items():
                self._keymap[key] = value

        def apply_quotation(self, mapping_dict: dict) -> None:
            for key, value in mapping_dict.items():
                self._keymap[key] = value, value, "Left"

        def ignore_capslock(self) -> None:
            for stat in self.key_status:
                for mod_key in self.mod_keys:
                    self._keymap[mod_key + stat + "Capslock"] = lambda: None

        def ignore_kanakey(self) -> None:
            for stat in self.key_status:
                for mod_key in self.mod_keys:
                    for vk in list(range(124, 136)) + list(range(240, 243)) + list(range(245, 254)):
                        self._keymap[mod_key + stat + str(vk)] = lambda: None

    GLOBAL_KEY_ALLOCATOR = KeyAllocator(keymap_global)
    GLOBAL_KEY_ALLOCATOR.cursor_keys()
    GLOBAL_KEY_ALLOCATOR.apply(
        {
            # close
            "LC-Q": ("A-F4"),
            # delete 2
            "LS-LC-U0-B": ("Back",) * 2,
            "LS-LC-U0-D": ("Delete",) * 2,
            # delete to bol / eol
            "S-U0-B": ("S-Home", "Delete"),
            "S-U0-D": ("S-End", "Delete"),
            # escape
            "O-(235)": ("Esc"),
            "U0-X": ("Esc"),
            # line selection
            "U1-A": ("End", "S-Home"),
            # punctuation
            "U0-Enter": ("Period"),
            "LS-U0-Enter": ("Comma"),
            "LC-U0-Enter": ("Slash"),
            "U0-U": "S-BackSlash",
            "U0-Z": ("Minus"),
            "U1-S": ("Slash"),
            "LS-U0-P": ("LS-Slash"),
            "LC-U0-P": ("LS-1"),
            # emacs-like backchar
            "LC-H": ("Back"),
            # Insert line
            "U0-I": ("End", "Enter"),
            "S-U0-I": ("Home", "Enter", "Up"),
            # Context menu
            "U0-C": ("Apps"),
            "S-U0-C": ("S-Apps"),
            # rename
            "U0-N": ("F2"),
            "LC-U0-N": ("F2"),
            # print
            "F1": ("C-P"),
            "U1-F1": ("F1"),
        }
    )

    GLOBAL_KEY_ALLOCATOR.apply_quotation(
        {
            "U0-2": "LS-2",
            "U0-7": "LS-7",
            "U0-AtMark": "LS-AtMark",
        }
    )

    ################################
    # functions for custom hotkey
    ################################

    def delay(msec: int = 50) -> None:
        time.sleep(msec / 1000)

    class Tap(NamedTuple):
        send: str
        has_real_key: bool

    class Taps:
        acceptable = (
            list(KeyCondition.str_vk_table_common)
            + list(KeyCondition.str_vk_table_std)
            + list(KeyCondition.str_vk_table_jpn)
        )

        def __init__(self) -> None:
            pass

        @classmethod
        def check(cls, s: str) -> bool:
            k = s.split("-")[-1].upper()
            return k in cls.acceptable

        @classmethod
        def from_sequence(cls, sequence: tuple) -> List[Tap]:
            seq = []
            for elem in sequence:
                key = Tap(elem, cls.check(elem))
                seq.append(key)
            return seq

    class VirtualFinger:
        def __init__(self, keymap: Keymap, inter_stroke_pause: int = 10) -> None:
            self._keymap = keymap
            self._inter_stroke_pause = inter_stroke_pause

        def _prepare(self) -> None:
            self._keymap.setInput_Modifier(0)
            self._keymap.beginInput()

        def _finish(self) -> None:
            self._keymap.endInput()

        def tap_keys(self, *keys) -> None:
            self._prepare()
            for key in keys:
                delay(self._inter_stroke_pause)
                self._keymap.setInput_FromString(str(key))
            self._finish()

        def tap_text(self, s: str) -> None:
            self._prepare()
            for c in str(s):
                delay(self._inter_stroke_pause)
                self._keymap.input_seq.append(pyauto.Char(c))
            self._finish()

        def tap_sequence(self, taps: List[Tap]) -> None:
            for tap in taps:
                if tap.has_real_key:
                    self.tap_keys(tap.send)
                else:
                    self.tap_text(tap.send)

    VIRTUAL_FINGER = VirtualFinger(keymap, 10)
    VIRTUAL_FINGER_QUICK = VirtualFinger(keymap, 0)

    class SKKKey:
        kata_key = "Q"
        kana_key = "C-J"
        latin_key = "S-L"
        cancel_key = "Esc"
        reconv_key = "LWin-Slash"
        abbrev_key = "Slash"
        convpoint_key = "S-0"

    class ImeControl(SKKKey):
        def __init__(self, keymap: Keymap, inter_stroke_pause: int = 10) -> None:
            self._keymap = keymap
            self._finger = VirtualFinger(self._keymap, inter_stroke_pause)

        def get_status(self) -> int:
            return self._keymap.getWindow().getImeStatus()

        def set_status(self, mode: int) -> None:
            if self.get_status() != mode:
                self._keymap.getWindow().setImeStatus(mode)

        def is_enabled(self) -> bool:
            return self.get_status() == 1

        def enable(self) -> None:
            self.set_status(1)

        def enable_skk(self) -> None:
            self.enable()
            self._finger.tap_keys(self.kana_key)

        def to_skk_latin(self) -> None:
            self.enable_skk()
            self._finger.tap_keys(self.latin_key)

        def to_skk_abbrev(self) -> None:
            self.enable_skk()
            self._finger.tap_keys(self.abbrev_key)

        def to_skk_kata(self) -> None:
            self.enable_skk()
            self._finger.tap_keys(self.kata_key)

        def start_skk_conv(self) -> None:
            self.enable_skk()
            self._finger.tap_keys(self.convpoint_key)

        def reconvert_with_skk(self) -> None:
            self.enable_skk()
            self._finger.tap_keys(self.reconv_key, self.cancel_key)

        def disable(self) -> None:
            self.set_status(0)

    IME_CONTROL = ImeControl(keymap)

    class ClipHandler:
        @staticmethod
        def get_string() -> str:
            return ckit.getClipboardText() or ""

        @staticmethod
        def set_string(s: str) -> None:
            ckit.setClipboardText(str(s))

        @classmethod
        def paste(cls, s: str, format_func: Union[Callable, None] = None) -> None:
            if format_func is not None:
                cls.set_string(format_func(s))
            else:
                cls.set_string(s)
            VIRTUAL_FINGER.tap_keys("C-V")

        @classmethod
        def paste_current(cls, format_func: Union[Callable, None] = None) -> None:
            cls.paste(cls.get_string(), format_func)

        @classmethod
        def after_copy(cls, deferred: Callable) -> None:
            cb = cls.get_string()
            VIRTUAL_FINGER.tap_keys("C-C")

            def _watch_clipboard(job_item: ckit.JobItem) -> None:
                job_item.origin = cb
                job_item.copied = ""
                interval = 10
                timeout = interval * 20
                while timeout > 0:
                    delay(interval)
                    s = cls.get_string()
                    if 0 < len(s.strip()) and s != job_item.origin:
                        job_item.copied = s
                        return
                    timeout -= interval

            subthread_run(_watch_clipboard, deferred)

        @classmethod
        def append(cls) -> None:

            def _push(job_item: ckit.JobItem) -> None:
                cls.set_string(job_item.origin + "\n" + job_item.copied)

            cls.after_copy(_push)

    class LazyKeymap:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        def defer(self, func: Callable, msec: int = 20) -> Callable:
            def _wrapper() -> None:
                self._keymap.delayedCall(func, msec)

            return _wrapper

    LAZY_KEYMAP = LazyKeymap(keymap)

    class KeyPuncher:
        def __init__(
            self,
            keymap: Keymap,
            recover_ime: bool = False,
            inter_stroke_pause: int = 0,
            defer_msec: int = 0,
        ) -> None:
            self._recover_ime = recover_ime
            self._defer_msec = defer_msec
            self._keymap = keymap
            self._finger = VirtualFinger(self._keymap, inter_stroke_pause)
            self._control = ImeControl(self._keymap)
            self._lazy_keymap = LazyKeymap(self._keymap)

        def invoke(self, *sequence) -> Callable:
            seq = Taps().from_sequence(sequence)

            def _input() -> None:
                self._control.disable()
                self._finger.tap_sequence(seq)
                if self._recover_ime:
                    self._control.enable()

            def __inhook_executer() -> None:
                self._keymap.hookCall(_input)

            return self._lazy_keymap.defer(__inhook_executer, self._defer_msec)

    MILD_PUNCHER = KeyPuncher(keymap, defer_msec=20)
    GENTLE_PUNCHER = KeyPuncher(keymap, defer_msec=50)

    keymap_global["U0-4"] = GENTLE_PUNCHER.invoke("$_")

    ################################
    # custom hotkey
    ################################

    keymap_global["LC-U0-C"] = ClipHandler().append

    # ime: Japanese / Foreign
    keymap_global["U1-J"] = IME_CONTROL.enable_skk
    keymap_global["LC-U0-I"] = IME_CONTROL.to_skk_kata
    keymap_global["U0-F"] = IME_CONTROL.disable
    keymap_global["LS-U0-F"] = IME_CONTROL.enable_skk
    keymap_global["S-U1-J"] = IME_CONTROL.to_skk_latin
    keymap_global["U1-I"] = IME_CONTROL.reconvert_with_skk
    keymap_global["LS-U1-I"] = IME_CONTROL.reconv_key
    keymap_global["O-(236)"] = IME_CONTROL.to_skk_abbrev
    keymap_global["LS-(236)"] = IME_CONTROL.start_skk_conv

    # paste as plaintext
    keymap_global["U0-V"] = LAZY_KEYMAP.defer(ClipHandler().paste_current)

    # paste as plaintext (with trimming removable whitespaces)
    class StrCleaner:
        @staticmethod
        def clear_space(s: str) -> str:
            return s.strip().translate(
                str.maketrans(
                    "",
                    "",
                    "\u0009\u0020\u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c\u200d\u200e\u200f\u202f\u205f\u3000\ufeff",
                )
            )

        @classmethod
        def invoke(cls, remove_white: bool = False, include_linebreak: bool = False) -> Callable:
            def _cleaner(s: str) -> str:
                s = s.strip()
                if remove_white:
                    s = cls.clear_space(s)
                if include_linebreak:
                    s = "".join(s.splitlines())
                return s

            def _paster() -> None:
                ClipHandler().paste_current(_cleaner)

            return _paster

        @classmethod
        def apply(cls, km: WindowKeymap, custom_key: str) -> None:
            for mod_ctrl, remove_white in {
                "": False,
                "LC-": True,
            }.items():
                for mod_shift, include_linebreak in {
                    "": False,
                    "LS-": True,
                }.items():
                    km[mod_ctrl + mod_shift + custom_key] = cls.invoke(
                        remove_white, include_linebreak
                    )

    StrCleaner().apply(keymap_global, "U1-V")

    # paste with quote mark
    def paste_with_anchor(join_lines: bool = False) -> Callable:
        def _formatter(s: str) -> str:
            lines = s.strip().splitlines()
            if join_lines:
                return "> " + "".join([line.strip() for line in lines])
            return "\n".join(["> " + line for line in lines])

        def _paster() -> None:
            ClipHandler().paste_current(_formatter)

        return _paster

    keymap_global["U1-Q"] = paste_with_anchor(False)
    keymap_global["C-U1-Q"] = paste_with_anchor(True)

    # open url in browser
    def open_selected_url():
        def _open(job_item: ckit.JobItem) -> None:
            if job_item.copied:
                u = job_item.copied
            else:
                u = job_item.origin
            shell_exec(u.strip())

        ClipHandler().after_copy(_open)

    keymap_global["C-U0-O"] = open_selected_url

    ################################
    # config menu
    ################################

    class ConfigMenu:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        @staticmethod
        def paste_config() -> None:
            s = Path(ckit.dataPath(), "config.py").read_text("utf-8")
            ClipHandler().paste(s)

        def reload_config(self) -> None:
            ckit.JobQueue.cancelAll()
            self._keymap.configure()
            self._keymap.updateKeymap()
            # self._keymap.console_window.reloadTheme()
            ts = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
            balloon("{} reloaded config.py".format(ts))

        @staticmethod
        def open_dir(path: str) -> None:
            if not smart_check_path(path):
                balloon("config not found: {}".format(path))
                return
            dir_path = path
            if (real_path := os.path.realpath(path)) != dir_path:
                dir_path = os.path.dirname(real_path)
            handler = PathHandler(dir_path)
            if handler.is_accessible():
                if KEYHAC_EDITOR == "notepad.exe":
                    balloon("keyhac editor 'notepad.exe' cannot open directory.")
                    handler.run()
                else:
                    PathHandler(KEYHAC_EDITOR).run(handler.path)
            else:
                balloon("cannot find path: '{}'".format(handler.path))

        def open_keyhac_repo(self) -> None:
            config_path = os.path.join(os.environ.get("APPDATA"), "Keyhac")
            self.open_dir(config_path)

        def open_skk_repo(self) -> None:
            config_path = os.path.join(os.environ.get("APPDATA"), "CorvusSKK")
            self.open_dir(config_path)

        @staticmethod
        def open_skk_config() -> None:
            skk_path = PathHandler(r"C:\Windows\System32\IME\IMCRVSKK\imcrvcnf.exe")
            skk_path.run()

        def apply(self, km: WindowKeymap) -> None:
            for key, func in {
                "R": self.reload_config,
                "E": self.open_keyhac_repo,
                "C-E": self.open_skk_repo,
                "P": self.paste_config,
                "S": self.open_skk_config,
                "X": lambda: None,
            }.items():
                km[key] = LAZY_KEYMAP.defer(func, 50)

    CONFIG_MENU = ConfigMenu(keymap)
    keymap_global["LC-U0-X"] = keymap.defineMultiStrokeKeymap()
    CONFIG_MENU.apply(keymap_global["LC-U0-X"])

    keymap.editor = lambda _: CONFIG_MENU.open_keyhac_repo()

    keymap_global["U1-F12"] = LAZY_KEYMAP.defer(ConfigMenu(keymap).reload_config, 50)

    ################################
    # class for position on monitor
    ################################

    class Rect(NamedTuple):
        left: int
        top: int
        right: int
        bottom: int

    class WndRect:
        def __init__(self, keymap: Keymap, rect: Rect) -> None:
            self._keymap = keymap
            self._rect = rect

        def check_rect(self, wnd: pyauto.Window) -> bool:
            return wnd.getRect() == self._rect

        def snap(self) -> None:

            def _snap(_) -> None:
                wnd = self._keymap.getTopLevelWindow()
                if not wnd:
                    return
                if CheckWnd.is_keyhac_console(wnd):
                    return
                if self.check_rect(wnd):
                    wnd.maximize()
                    return
                if wnd.isMaximized():
                    wnd.restore()
                    delay()
                trial_limit = 2
                while trial_limit > 0:
                    wnd.setRect(self._rect)
                    delay()
                    if self.check_rect(wnd):
                        return
                    trial_limit -= 1

            subthread_run(_snap)

        def get_vertical_half(self, upper: bool) -> list:
            half_height = int((self._rect.bottom + self._rect.top) / 2)
            r = list(self._rect)
            if upper:
                r[3] = half_height
            else:
                r[1] = half_height
            if 200 < abs(r[1] - r[3]):
                return r
            return []

        def get_horizontal_half(self, leftward: bool) -> list:
            half_width = int((self._rect.right + self._rect.left) / 2)
            r = list(self._rect)
            if leftward:
                r[2] = half_width
            else:
                r[0] = half_width
            if 300 < abs(r[0] - r[2]):
                return r
            return []

    class MonitorRect:
        def __init__(self, keymap: Keymap, rect: Rect) -> None:
            self._keymap = keymap
            self.left, self.top, self.right, self.bottom = rect
            self.max_width = self.right - self.left
            self.max_height = self.bottom - self.top
            self.possible_width = self.get_variation(self.max_width)
            self.possible_height = self.get_variation(self.max_height)
            self.area_mapping: Dict[str, Dict[str, WndRect]] = {}

        @staticmethod
        def get_variation(max_size: int) -> dict:
            return {
                "small": int(max_size / 3),
                "middle": int(max_size / 2),
                "large": int(max_size * 2 / 3),
            }

        def set_center_rect(self) -> None:
            d = {}
            for size, px in self.possible_width.items():
                lx = self.left + int((self.max_width - px) / 2)
                wr = WndRect(self._keymap, Rect(lx, self.top, lx + px, self.bottom))
                d[size] = wr
            self.area_mapping["center"] = d

        def set_horizontal_rect(self) -> None:
            for pos in ("left", "right"):
                d = {}
                for size, px in self.possible_width.items():
                    if pos == "right":
                        lx = self.right - px
                    else:
                        lx = self.left
                    wr = WndRect(self._keymap, Rect(lx, self.top, lx + px, self.bottom))
                    d[size] = wr
                self.area_mapping[pos] = d

        def set_vertical_rect(self) -> None:
            for pos in ("top", "bottom"):
                d = {}
                for size, px in self.possible_height.items():
                    if pos == "bottom":
                        ty = self.bottom - px
                    else:
                        ty = self.top
                    wr = WndRect(self._keymap, Rect(self.left, ty, self.right, ty + px))
                    d[size] = wr
                self.area_mapping[pos] = d

    class CurrentMonitors:
        def __init__(self, keymap: Keymap) -> None:
            ms = []
            for mi in pyauto.Window.getMonitorInfo():
                mr = MonitorRect(keymap, Rect(*mi[1]))
                mr.set_center_rect()
                mr.set_horizontal_rect()
                mr.set_vertical_rect()
                if mi[2] == 1:  # main monitor
                    ms.insert(0, mr)
                else:
                    ms.append(mr)
            self._monitors = ms

        @property
        def monitors(self) -> list:
            return self._monitors

    ################################
    # set cursor position
    ################################

    class CursorPos:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap
            self.pos = []
            monitors = CurrentMonitors(keymap).monitors
            for monitor in monitors:
                for i in (1, 3):
                    y = monitor.top + int(monitor.max_height / 2)
                    x = monitor.left + int(monitor.max_width / 4) * i
                    self.pos.append([x, y])

        def get_position_index(self) -> int:
            x, y = pyauto.Input.getCursorPos()
            for i, p in enumerate(self.pos):
                if p[0] == x and p[1] == y:
                    return i
            return -1

        def snap(self) -> None:
            idx = self.get_position_index()
            if idx < 0 or idx == len(self.pos) - 1:
                self.set_position(*self.pos[0])
            else:
                self.set_position(*self.pos[idx + 1])

        def set_position(self, x: int, y: int) -> None:
            self._keymap.beginInput()
            self._keymap.input_seq.append(pyauto.MouseMove(x, y))
            self._keymap.endInput()

        def snap_to_center(self) -> None:
            wnd = self._keymap.getTopLevelWindow()
            wnd_left, wnd_top, wnd_right, wnd_bottom = wnd.getRect()
            to_x = int((wnd_left + wnd_right) / 2)
            to_y = int((wnd_bottom + wnd_top) / 2)
            self.set_position(to_x, to_y)

    CURSOR_POS = CursorPos(keymap)

    keymap_global["O-RCtrl"] = CURSOR_POS.snap
    keymap_global["O-RShift"] = CURSOR_POS.snap_to_center

    ################################
    # set window position
    ################################

    def apply_window_mover(km: WindowKeymap) -> None:
        for key, delta in {
            "Left": (-10, 0),
            "Right": (+10, 0),
            "Up": (0, -10),
            "Down": (0, +10),
        }.items():
            x, y = delta
            for mod, scale in {"": 15, "S-": 5, "C-": 5, "S-C-": 1}.items():
                km[mod + "U0-" + key] = keymap.MoveWindowCommand(x * scale, y * scale)

    apply_window_mover(keymap_global)

    keymap_global["U1-L"] = "LWin-Right"
    keymap_global["U1-H"] = "LWin-Left"

    keymap_global["U1-M"] = keymap.defineMultiStrokeKeymap()

    def maximize_window():
        def _maximize(_) -> None:
            delay()
            keymap.getTopLevelWindow().maximize()

        subthread_run(_maximize)

    keymap_global["U1-M"]["X"] = maximize_window

    class WndPosAllocator:
        monitor_dict = {
            "": 0,
            "A-": 1,
        }
        size_dict = {
            "": "middle",
            "S-": "large",
            "C-": "small",
        }
        snap_key_dict = {
            "H": "left",
            "L": "right",
            "J": "bottom",
            "K": "top",
            "Left": "left",
            "Right": "right",
            "Down": "bottom",
            "Up": "top",
            "M": "center",
        }

        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        def alloc_flexible(self, km: WindowKeymap) -> None:
            monitors = CurrentMonitors(self._keymap).monitors
            for mod_mntr, mntr_idx in self.monitor_dict.items():
                for mod_area, size in self.size_dict.items():
                    for key, pos in self.snap_key_dict.items():
                        if mntr_idx < len(monitors):
                            wnd_rect = monitors[mntr_idx].area_mapping[pos][size]
                            km[mod_mntr + mod_area + key] = LAZY_KEYMAP.defer(wnd_rect.snap, 50)

        def alloc_maximize(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, towards in mapping_dict.items():

                def _snap() -> None:
                    def _maximize(_) -> None:
                        self._keymap.getTopLevelWindow().maximize()

                    def _snapper(_) -> None:
                        VIRTUAL_FINGER.tap_keys("LShift-LWin-" + towards)

                    subthread_run(_maximize, _snapper)

                km[key] = _snap

    WND_POS_ALLOCATOR = WndPosAllocator(keymap)
    WND_POS_ALLOCATOR.alloc_flexible(keymap_global["U1-M"])

    WND_POS_ALLOCATOR.alloc_maximize(keymap_global, {"LC-U1-L": "Right", "LC-U1-H": "Left"})
    WND_POS_ALLOCATOR.alloc_maximize(
        keymap_global["U1-M"],
        {
            "U0-L": "Right",
            "U0-J": "Right",
            "U0-H": "Left",
            "U0-K": "Left",
        },
    )

    class WndShrinker:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        def invoke_snapper(self, horizontal: bool, default_pos: bool) -> Callable:
            def _snapper() -> None:
                def _snap(_) -> None:
                    wnd = self._keymap.getTopLevelWindow()
                    wr = WndRect(self._keymap, Rect(*wnd.getRect()))
                    rect = []
                    if horizontal:
                        rect = wr.get_horizontal_half(default_pos)
                    else:
                        rect = wr.get_vertical_half(default_pos)
                    if len(rect):
                        if wnd.isMaximized():
                            wnd.restore()
                            delay()
                        wnd.setRect(rect)

                subthread_run(_snap)

            return LAZY_KEYMAP.defer(_snapper, 50)

        def apply(self, km: WindowKeymap) -> None:
            for key, params in {
                "H": {"horizontal": True, "default_pos": True},
                "L": {"horizontal": True, "default_pos": False},
                "K": {"horizontal": False, "default_pos": True},
                "J": {"horizontal": False, "default_pos": False},
            }.items():
                km["U1-" + key] = self.invoke_snapper(**params)

    WndShrinker(keymap).apply(keymap_global["U1-M"])

    ################################
    # input customize
    ################################

    class SKKSender:
        def __init__(
            self,
            keymap: Keymap,
            inter_stroke_pause: int = 0,
        ) -> None:
            self._finger = VirtualFinger(keymap, inter_stroke_pause)
            self._control = ImeControl(keymap)

        def under_kanamode(self, *sequence) -> Callable:
            taps = Taps().from_sequence(sequence)

            def _send() -> None:
                self._control.enable_skk()
                self._finger.tap_sequence(taps)

            return _send

        def under_latinmode(self, *sequence) -> Callable:
            taps = Taps().from_sequence(sequence)

            def _send() -> None:
                self._control.to_skk_latin()
                self._finger.tap_sequence(taps)

            return _send

    # select-to-left with ime control
    keymap_global["U1-B"] = SKKSender(keymap).under_kanamode("S-Left")
    keymap_global["LS-U1-B"] = SKKSender(keymap).under_kanamode("S-Right")
    keymap_global["U1-Space"] = SKKSender(keymap).under_kanamode("C-S-Left")
    keymap_global["U1-N"] = SKKSender(keymap).under_kanamode(
        "C-S-Left", ImeControl.convpoint_key, "S-4"
    )
    keymap_global["U1-4"] = SKKSender(keymap).under_kanamode(SKKKey.convpoint_key, "S-4")

    class SKKFinisher:
        def __init__(self, keymap: Keymap, to_kanamode: bool = True) -> None:
            self._skk_sender = SKKSender(keymap)
            self._to_kanamode = to_kanamode

        def invoke(self, *sequence) -> Callable:
            if self._to_kanamode:
                sequence = list(sequence) + [ImeControl.kana_key]
            return self._skk_sender.under_latinmode(*sequence)

        def apply(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, sent in mapping_dict.items():
                km[key] = self.invoke(sent)

        def apply_circumfix(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, circumfix in mapping_dict.items():
                _, suffix = circumfix
                sequence = circumfix + ["Left"] * len(suffix)
                km[key] = self.invoke(*sequence)

    # insert honorific
    def invoke_honorific_sender(km: WindowKeymap) -> None:
        finisher = SKKFinisher(keymap)
        for key, hono in {"U0": "先生", "U1": "様"}.items():
            for mod, suffix in {"": "", "C-": "方"}.items():
                km[mod + key + "-Tab"] = finisher.invoke(hono + suffix)

    invoke_honorific_sender(keymap_global)

    # markdown list
    keymap_global["S-U0-8"] = SKKFinisher(keymap).invoke("- ")
    keymap_global["U1-1"] = SKKFinisher(keymap).invoke("1. ")

    SKKFinisher(keymap).apply(
        keymap_global,
        {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Minus": "\u3000\u2015\u2015",
            "S-U0-Period": "\uff0e",  # FULLWIDTH FULL STOP
            "U0-Minus": "\u2015\u2015",  # HORIZONTAL BAR * 2
        },
    )

    SKKFinisher(keymap).apply_circumfix(
        keymap_global,
        {
            "U0-8": ["\u300e", "\u300f"],  # WHITE CORNER BRACKET 『』
            "U0-9": ["\u3010", "\u3011"],  # BLACK LENTICULAR BRACKET 【】
            "U0-OpenBracket": ["\u300c", "\u300d"],  # CORNER BRACKET 「」
            "U1-2": ["\u201c", "\u201d"],  # DOUBLE QUOTATION MARK “”
            "U1-7": ["\u2018", "\u2019"],  # SINGLE QUOTATION MARK ‘’
            "U1-8": ["\uff08", "\uff09"],  # FULLWIDTH PARENTHESIS （）
            "U1-Atmark": ["\uff08`", "`\uff09"],  # FULLWIDTH PARENTHESIS and BackTick （``）
            "U0-Y": ["\u3008", "\u3009"],  # ANGLE BRACKET
            "S-U0-Y": ["\u300a", "\u300b"],  # DOUBLE ANGLE BRACKET
            "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
        },
    )

    SKKFinisher(keymap, False).apply(
        keymap_global,
        {
            "U0-1": "S-1",
            "U0-Colon": "Colon",
            "U0-Slash": "Slash",
            "U1-Minus": "Minus",
            "U0-SemiColon": "SemiColon",
            "U1-SemiColon": "+:",
            "U0-Comma": "Comma",
            "U0-Period": "Period",
        },
    )
    SKKFinisher(keymap, False).apply_circumfix(
        keymap_global,
        {
            "U0-CloseBracket": ["[", "]"],
            "U1-9": ["(", ")"],
            "U1-CloseBracket": ["{", "}"],
            "U0-Caret": ["~~", "~~"],
            "U0-T": ["<", ">"],
            "LS-U0-T": ["</", ">"],
        },
    )

    ################################
    # web search
    ################################

    class KangxiRadicals:
        mapping = {
            "\u2f00": "\u4e00",
            "\u2f01": "\u4e28",
            "\u2f02": "\u4e36",
            "\u2f03": "\u4e3f",
            "\u2f04": "\u4e59",
            "\u2f05": "\u4e85",
            "\u2f06": "\u4e8c",
            "\u2f07": "\u4ea0",
            "\u2f08": "\u4eba",
            "\u2f09": "\u513f",
            "\u2f0a": "\u5165",
            "\u2f0b": "\u516b",
            "\u2f0c": "\u5182",
            "\u2f0d": "\u5196",
            "\u2f0e": "\u51ab",
            "\u2f0f": "\u51e0",
            "\u2f10": "\u51f5",
            "\u2f11": "\u5200",
            "\u2f12": "\u529b",
            "\u2f13": "\u52f9",
            "\u2f14": "\u5315",
            "\u2f15": "\u531a",
            "\u2f16": "\u5338",
            "\u2f17": "\u5341",
            "\u2f18": "\u535c",
            "\u2f19": "\u5369",
            "\u2f1a": "\u5382",
            "\u2f1b": "\u53b6",
            "\u2f1c": "\u53c8",
            "\u2f1d": "\u53e3",
            "\u2f1e": "\u56d7",
            "\u2f1f": "\u571f",
            "\u2f20": "\u58eb",
            "\u2f21": "\u5902",
            "\u2f22": "\u590a",
            "\u2f23": "\u5915",
            "\u2f24": "\u5927",
            "\u2f25": "\u5973",
            "\u2f26": "\u5b50",
            "\u2f27": "\u5b80",
            "\u2f28": "\u5bf8",
            "\u2f29": "\u5c0f",
            "\u2f2a": "\u5c22",
            "\u2f2b": "\u5c38",
            "\u2f2c": "\u5c6e",
            "\u2f2d": "\u5c71",
            "\u2f2e": "\u5ddb",
            "\u2f2f": "\u5de5",
            "\u2f30": "\u5df1",
            "\u2f31": "\u5dfe",
            "\u2f32": "\u5e72",
            "\u2f33": "\u5e7a",
            "\u2f34": "\u5e7f",
            "\u2f35": "\u5ef4",
            "\u2f36": "\u5efe",
            "\u2f37": "\u5f0b",
            "\u2f38": "\u5f13",
            "\u2f39": "\u5f50",
            "\u2f3a": "\u5f61",
            "\u2f3b": "\u5f73",
            "\u2f3c": "\u5fc3",
            "\u2f3d": "\u6208",
            "\u2f3e": "\u6238",
            "\u2f3f": "\u624b",
            "\u2f40": "\u652f",
            "\u2f41": "\u6534",
            "\u2f42": "\u6587",
            "\u2f43": "\u6597",
            "\u2f44": "\u65a4",
            "\u2f45": "\u65b9",
            "\u2f46": "\u65e0",
            "\u2f47": "\u65e5",
            "\u2f48": "\u66f0",
            "\u2f49": "\u6708",
            "\u2f4a": "\u6728",
            "\u2f4b": "\u6b20",
            "\u2f4c": "\u6b62",
            "\u2f4d": "\u6b79",
            "\u2f4e": "\u6bb3",
            "\u2f4f": "\u6bcb",
            "\u2f50": "\u6bd4",
            "\u2f51": "\u6bdb",
            "\u2f52": "\u6c0f",
            "\u2f53": "\u6c14",
            "\u2f54": "\u6c34",
            "\u2f55": "\u706b",
            "\u2f56": "\u722a",
            "\u2f57": "\u7236",
            "\u2f58": "\u723b",
            "\u2f59": "\u723f",
            "\u2f5a": "\u7247",
            "\u2f5b": "\u7259",
            "\u2f5c": "\u725b",
            "\u2f5d": "\u72ac",
            "\u2f5e": "\u7384",
            "\u2f5f": "\u7389",
            "\u2f60": "\u74dc",
            "\u2f61": "\u74e6",
            "\u2f62": "\u7518",
            "\u2f63": "\u751f",
            "\u2f64": "\u7528",
            "\u2f65": "\u7530",
            "\u2f66": "\u758b",
            "\u2f67": "\u7592",
            "\u2f68": "\u7676",
            "\u2f69": "\u767d",
            "\u2f6a": "\u76ae",
            "\u2f6b": "\u76bf",
            "\u2f6c": "\u76ee",
            "\u2f6d": "\u77db",
            "\u2f6e": "\u77e2",
            "\u2f6f": "\u77f3",
            "\u2f70": "\u793a",
            "\u2f71": "\u79b8",
            "\u2f72": "\u79be",
            "\u2f73": "\u7a74",
            "\u2f74": "\u7acb",
            "\u2f75": "\u7af9",
            "\u2f76": "\u7c73",
            "\u2f77": "\u7cf8",
            "\u2f78": "\u7f36",
            "\u2f79": "\u7f51",
            "\u2f7a": "\u7f8a",
            "\u2f7b": "\u7fbd",
            "\u2f7c": "\u8001",
            "\u2f7d": "\u800c",
            "\u2f7e": "\u8012",
            "\u2f7f": "\u8033",
            "\u2f80": "\u807f",
            "\u2f81": "\u8089",
            "\u2f82": "\u81e3",
            "\u2f83": "\u81ea",
            "\u2f84": "\u81f3",
            "\u2f85": "\u81fc",
            "\u2f86": "\u820c",
            "\u2f87": "\u821b",
            "\u2f88": "\u821f",
            "\u2f89": "\u826e",
            "\u2f8a": "\u8272",
            "\u2f8b": "\u8278",
            "\u2f8c": "\u864d",
            "\u2f8d": "\u866b",
            "\u2f8e": "\u8840",
            "\u2f8f": "\u884c",
            "\u2f90": "\u8863",
            "\u2f91": "\u897e",
            "\u2f92": "\u898b",
            "\u2f93": "\u89d2",
            "\u2f94": "\u8a00",
            "\u2f95": "\u8c37",
            "\u2f96": "\u8c46",
            "\u2f97": "\u8c55",
            "\u2f98": "\u8c78",
            "\u2f99": "\u8c9d",
            "\u2f9a": "\u8d64",
            "\u2f9b": "\u8d70",
            "\u2f9c": "\u8db3",
            "\u2f9d": "\u8eab",
            "\u2f9e": "\u8eca",
            "\u2f9f": "\u8f9b",
            "\u2fa0": "\u8fb0",
            "\u2fa1": "\u8fb5",
            "\u2fa2": "\u9091",
            "\u2fa3": "\u9149",
            "\u2fa4": "\u91c6",
            "\u2fa5": "\u91cc",
            "\u2fa6": "\u91d1",
            "\u2fa7": "\u9577",
            "\u2fa8": "\u9580",
            "\u2fa9": "\u961c",
            "\u2faa": "\u96b6",
            "\u2fab": "\u96b9",
            "\u2fac": "\u96e8",
            "\u2fad": "\u9751",
            "\u2fae": "\u975e",
            "\u2faf": "\u9762",
            "\u2fb0": "\u9769",
            "\u2fb1": "\u97cb",
            "\u2fb2": "\u97ed",
            "\u2fb3": "\u97f3",
            "\u2fb4": "\u9801",
            "\u2fb5": "\u98a8",
            "\u2fb6": "\u98db",
            "\u2fb7": "\u98df",
            "\u2fb8": "\u9996",
            "\u2fb9": "\u9999",
            "\u2fba": "\u99ac",
            "\u2fbb": "\u9aa8",
            "\u2fbc": "\u9ad8",
            "\u2fbd": "\u9adf",
            "\u2fbe": "\u9b25",
            "\u2fbf": "\u9b2f",
            "\u2fc0": "\u9b32",
            "\u2fc1": "\u9b3c",
            "\u2fc2": "\u9b5a",
            "\u2fc3": "\u9ce5",
            "\u2fc4": "\u9e75",
            "\u2fc5": "\u9e7f",
            "\u2fc6": "\u9ea5",
            "\u2fc7": "\u9ebb",
            "\u2fc8": "\u9ec3",
            "\u2fc9": "\u9ecd",
            "\u2fca": "\u9ed2",
            "\u2fcb": "\u9ef9",
            "\u2fcc": "\u9efd",
            "\u2fcd": "\u9f0e",
            "\u2fce": "\u9f13",
            "\u2fcf": "\u9f20",
            "\u2fd0": "\u9f3b",
            "\u2fd1": "\u9f4a",
            "\u2fd2": "\u9f52",
            "\u2fd3": "\u9f8d",
            "\u2fd4": "\u9f9c",
            "\u2fd5": "\u9fa0",
        }

        @classmethod
        def fix(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.mapping))

    class UnicodeMapper:
        def __init__(self, repl: str) -> None:
            if len(repl):
                self._repl = ord(repl)
            else:
                self._repl = None
            self._mapping = {}

        def register(self, char_code: int) -> None:
            self._mapping[char_code] = self._repl

        def register_range(self, pair: list) -> None:
            start, end = pair
            for i in range(int(start, 16), int(end, 16) + 1):
                self.register(i)

        def register_ranges(self, pairs: list) -> None:
            for pair in pairs:
                self.register_range(pair)

        @property
        def mapping(self) -> dict:
            return self._mapping

    class SearchNoiseMapping:
        def __init__(self, repl: str) -> None:
            _mapper = UnicodeMapper(repl)
            _mapper.register(int("30FB", 16))  # KATAKANA MIDDLE DOT
            _mapper.register_range(["2018", "201F"])  # quotation
            _mapper.register_range(["2E80", "2EF3"])  # kangxi
            _mapper.register_ranges(
                [  # ascii
                    ["0021", "002F"],
                    ["003A", "0040"],
                    ["005B", "0060"],
                    ["007B", "007E"],
                ]
            )
            _mapper.register_ranges(
                [  # bars
                    ["2010", "2017"],
                    ["2500", "2501"],
                    ["2E3A", "2E3B"],
                ]
            )
            _mapper.register_ranges(
                [  # fullwidth
                    ["25A0", "25EF"],
                    ["3000", "3004"],
                    ["3008", "3040"],
                    ["3097", "30A0"],
                    ["3097", "30A0"],
                    ["30FD", "30FF"],
                    ["FF01", "FF0F"],
                    ["FF1A", "FF20"],
                    ["FF3B", "FF40"],
                    ["FF5B", "FF65"],
                ]
            )

            self._mapping = _mapper.mapping

        def cleanup(self, s: str) -> str:
            return s.translate(str.maketrans(self._mapping))

    SEARCH_NOISE_MAPPIING = SearchNoiseMapping(" ")

    class SearchQuery:
        def __init__(self, query: str) -> None:
            self._query = ""
            lines = (
                query.strip()
                .replace("\u200b", "")
                .replace("\u3000", " ")
                .replace("\t", " ")
                .splitlines()
            )
            for line in lines:
                self._query += self.format_line(line)

        @staticmethod
        def format_line(s: str) -> str:
            if s.endswith("-"):
                return s.rstrip("-")
            if len(s.strip()):
                if s[-1].encode("utf-8").isalnum():
                    return s + " "
                return s.rstrip()
            return ""

        def fix_kangxi(self) -> None:
            self._query = KangxiRadicals().fix(self._query)

        def remove_honorific(self) -> None:
            for honor in ["先生", "様"]:
                self._query = self._query.replace(honor, " ")

        def remove_editorial_style(self) -> None:
            for honor in ["監修", "共著", "編著", "共編著", "共編", "分担執筆", "et al."]:
                self._query = self._query.replace(honor, " ")

        def remove_hiragana(self) -> None:
            self._query = re.sub(r"[\u3041-\u3093]", " ", self._query)

        def encode(self, strict: bool = False) -> str:
            words = []
            for word in SEARCH_NOISE_MAPPIING.cleanup(self._query).split(" "):
                if len(word):
                    if strict:
                        words.append('"{}"'.format(word))
                    else:
                        words.append(word)
            return urllib.parse.quote(" ".join(words))

    class WebSearcher:
        def __init__(self, uri_mapping: dict) -> None:
            self._keymap = keymap
            self._uri_mapping = uri_mapping

        @staticmethod
        def invoke(uri: str, strict: bool = False, strip_hiragana: bool = False) -> Callable:
            def _searcher() -> None:
                def _search(job_item: ckit.JobItem) -> None:
                    s = job_item.copied
                    if len(s) < 1:
                        s = job_item.origin
                    query = SearchQuery(s)
                    query.fix_kangxi()
                    query.remove_honorific()
                    query.remove_editorial_style()
                    if strip_hiragana:
                        query.remove_hiragana()
                    shell_exec(uri.format(query.encode(strict)))

                ClipHandler().after_copy(_search)

            return LAZY_KEYMAP.defer(_searcher)

        @staticmethod
        def _mod_key(s: str) -> list:
            return ["", s + "-"]

        def apply(self, km: WindowKeymap) -> None:
            for shift_key in self._mod_key("S"):
                for ctrl_key in self._mod_key("C"):
                    is_strict = 0 < len(shift_key)
                    strip_hiragana = 0 < len(ctrl_key)
                    trigger_key = shift_key + ctrl_key + "U0-S"
                    km[trigger_key] = self._keymap.defineMultiStrokeKeymap()
                    for key, uri in self._uri_mapping.items():
                        km[trigger_key][key] = self.invoke(uri, is_strict, strip_hiragana)

    WebSearcher(
        {
            "A": "https://www.amazon.co.jp/s?i=stripbooks&k={}",
            "B": "https://www.google.com/search?nfpr=1&q=site%3Abooks.or.jp%20{}",
            "C": "https://ci.nii.ac.jp/books/search?q={}",
            "D": "https://duckduckgo.com/?q={}",
            "G": "http://www.google.com/search?nfpr=1&q={}",
            "H": "https://www.hanmoto.com/bd/search/order/desc/title/{}",
            "I": "https://www.google.com/search?udm=2&nfpr=1&q={}",
            "J": "https://eow.alc.co.jp/search?q={}",
            "K": "https://www.kinokuniya.co.jp/disp/CSfDispListPage_001.jsp?qs=true&ptk=01&q={}",
            "M": "https://www.google.co.jp/maps/search/{}",
            "N": "https://iss.ndl.go.jp/books?any={}",
            "P": "https://wordpress.org/openverse/search/?q={}",
            "R": "https://researchmap.jp/researchers?q={}",
            "S": "https://scholar.google.com/scholar?nfpr=1&as_vis=1&q={}",
            "T": "https://twitter.com/search?q={}",
            "Y": "https://duckduckgo.com/?q=site%3Ayuhikaku.co.jp%20{}",
            "W": "https://www.worldcat.org/search?q={}",
        }
    ).apply(keymap_global)

    ################################
    # activate window
    ################################

    class SystemBrowser:
        register_path = (
            r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice"
        )
        prog_id = ""
        with OpenKey(HKEY_CURRENT_USER, register_path) as key:
            prog_id = str(QueryValueEx(key, "ProgId")[0])

        @classmethod
        def get_commandline(cls) -> str:
            register_path = r"{}\shell\open\command".format(cls.prog_id)
            with OpenKey(HKEY_CLASSES_ROOT, register_path) as key:
                return str(QueryValueEx(key, "")[0])

        @classmethod
        def get_exe_path(cls) -> str:
            return re.sub(r"(^.+\.exe)(.*)", r"\1", cls.get_commandline()).replace('"', "")

        @classmethod
        def get_exe_name(cls) -> str:
            return Path(cls.get_exe_path()).name

        @classmethod
        def get_wnd_class(cls) -> str:
            return {
                "chrome.exe": "Chrome_WidgetWin_1",
                "vivaldi.exe": "Chrome_WidgetWin_1",
                "firefox.exe": "MozillaWindowClass",
            }.get(cls.get_exe_name(), "Chrome_WidgetWin_1")

    DEFAULT_BROWSER = SystemBrowser()

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
            if not fnmatch.fnmatch(wnd.getProcessName(), self.exe_name):
                return True
            if self.class_name and not fnmatch.fnmatch(wnd.getClassName(), self.class_name):
                return True
            if len(wnd.getText()) < 1:
                return True
            popup = wnd.getLastActivePopup()
            if not popup:
                return True
            self.found = popup
            return False

    class PseudoCuteExec:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        @staticmethod
        def is_already_active(wnd: pyauto.Window) -> bool:
            return pyauto.Window.getForeground() == wnd

        def activate_wnd(self, target: pyauto.Window) -> bool:
            if self.is_already_active(target):
                return True
            if target.isMinimized():
                target.restore()
            interval = 20
            timeout = interval * 50
            while timeout > 0:
                try:
                    target.setForeground()
                    delay(interval)
                    if self.is_already_active(target):
                        target.setForeground(True)
                        return True
                except:
                    return False
                timeout -= interval
            return False

        def invoke(self, exe_name: str, class_name: str = "", exe_path: str = "") -> Callable:
            def _executer() -> None:

                def _activate(job_item: ckit.JobItem) -> None:
                    delay()
                    job_item.results = []
                    scanner = WndScanner(exe_name, class_name)
                    scanner.scan()
                    if scanner.found:
                        result = self.activate_wnd(scanner.found)
                        job_item.results.append(result)

                def _finished(job_item: ckit.JobItem) -> None:
                    if len(job_item.results) < 1:
                        if exe_path:
                            PathHandler(exe_path).run()
                        return
                    if not job_item.results[-1]:
                        VirtualFinger(self._keymap).tap_keys("LCtrl-LAlt-Tab")

                subthread_run(_activate, _finished)

            return _executer

        def apply(self, wnd_keymap: WindowKeymap, remap_table: dict = {}) -> None:
            for key, params in remap_table.items():
                func = self.invoke(*params)
                wnd_keymap[key] = LAZY_KEYMAP.defer(func, 80)

    PSEUDO_CUTEEXEC = PseudoCuteExec(keymap)

    PSEUDO_CUTEEXEC.apply(
        keymap_global,
        {
            "U1-F": (
                "cfiler.exe",
                "CfilerWindowClass",
                r"${USERPROFILE}\Sync\portable_app\cfiler\cfiler.exe",
            ),
            "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
            "LC-U1-M": (
                "Mery.exe",
                "TChildForm",
                r"${USERPROFILE}\AppData\Local\Programs\Mery\Mery.exe",
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
        },
    )

    keymap_global["U1-C"] = keymap.defineMultiStrokeKeymap()
    PSEUDO_CUTEEXEC.apply(
        keymap_global["U1-C"],
        {
            "Space": (
                DEFAULT_BROWSER.get_exe_name(),
                DEFAULT_BROWSER.get_wnd_class(),
                DEFAULT_BROWSER.get_exe_path(),
            ),
            "C": (
                "chrome.exe",
                "Chrome_WidgetWin_1",
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            ),
            "G": (
                "msedge.exe",
                "Chrome_WidgetWin_1",
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            ),
            "D": (
                "vivaldi.exe",
                "Chrome_WidgetWin_1",
                r"${USERPROFILE}\AppData\Local\Vivaldi\Application\vivaldi.exe",
            ),
            "S": (
                "slack.exe",
                "Chrome_WidgetWin_1",
                r"${USERPROFILE}\AppData\Local\slack\slack.exe",
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
            "V": ("Code.exe", "Chrome_WidgetWin_1", KEYHAC_EDITOR),
            "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1"),
            "M": (
                "Mery.exe",
                "TChildForm",
                r"${USERPROFILE}\AppData\Local\Programs\Mery\Mery.exe",
            ),
            "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
        },
    )

    def fuzzy_window_switcher() -> None:
        if not check_fzf():
            balloon("cannot find fzf on PC.")
            return

        executer = PseudoCuteExec(keymap)

        black_list = [
            "explorer.exe",
            "MouseGestureL.exe",
            "TextInputHost.exe",
            "SystemSettings.exe",
            "ApplicationFrameHost.exe",
        ]

        def _fzf_wnd(job_item: ckit.JobItem) -> None:
            job_item.result = []
            d = {}
            proc = subprocess.Popen(
                ["fzf.exe"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
            )

            def _walk(wnd: pyauto.Window, _) -> bool:
                if not wnd:
                    return False
                if not wnd.isVisible():
                    return True
                if not wnd.isEnabled():
                    return True
                if CheckWnd.is_keyhac_console(wnd):
                    return True
                if wnd.getProcessName() in black_list:
                    return True
                if len(wnd.getText()) < 1:
                    return True
                if popup := wnd.getLastActivePopup():
                    n = "{} [{}]".format(
                        popup.getProcessName().replace(".exe", ""), popup.getText()
                    )
                    d[n] = popup
                    proc.stdin.write(n + "\n")
                return True

            try:
                pyauto.Window.enum(_walk, None)
                proc.stdin.close()
            except Exception as e:
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
            found = d.get(result, None)
            if found:
                job_item.result.append(executer.activate_wnd(found))

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.isCanceled():
                return
            if 0 < len(job_item.result):
                if not job_item.result[0]:
                    VIRTUAL_FINGER.tap_keys("LCtrl-LAlt-Tab")

        subthread_run(_fzf_wnd, _finished)

    keymap_global["U1-E"] = LAZY_KEYMAP.defer(fuzzy_window_switcher, 120)

    def invoke_draft() -> None:
        def _invoke(_) -> None:
            PathHandler(r"${USERPROFILE}\Personal\draft.txt").run()

        subthread_run(_invoke)

    keymap_global["LS-LC-U1-M"] = invoke_draft

    def search_on_browser() -> None:
        if keymap.getWindow().getProcessName() == DEFAULT_BROWSER.get_exe_name():
            VIRTUAL_FINGER.tap_keys("C-T")
            return

        def _activate(job_item: ckit.JobItem) -> None:
            delay()
            job_item.results = []
            scanner = WndScanner(DEFAULT_BROWSER.get_exe_name(), DEFAULT_BROWSER.get_wnd_class())
            scanner.scan()
            if scanner.found:
                result = PSEUDO_CUTEEXEC.activate_wnd(scanner.found)
                job_item.results.append(result)

        def _finished(job_item: ckit.JobItem) -> None:
            if len(job_item.results) < 1:
                PathHandler(DEFAULT_BROWSER.get_exe_path()).run()
                return
            if not job_item.results[-1]:
                VIRTUAL_FINGER.tap_keys("LCtrl-LAlt-Tab")
                return
            VIRTUAL_FINGER.tap_keys("C-T")

        subthread_run(_activate, _finished)

    keymap_global["U0-Q"] = LAZY_KEYMAP.defer(search_on_browser, 10)

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=CheckWnd.is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"
    keymap_browser["LC-L"] = KeyPuncher(keymap, defer_msec=50, recover_ime=False).invoke("C-L")
    keymap_browser["LC-F"] = KeyPuncher(keymap, defer_msec=50, recover_ime=False).invoke("C-F")

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None

    # slack
    keymap_slack = keymap.defineWindowKeymap(exe_name="slack.exe", class_name="Chrome_WidgetWin_1")
    keymap_slack["C-S-K"] = GENTLE_PUNCHER.invoke("C-S-K")
    keymap_slack["F3"] = GENTLE_PUNCHER.invoke("C-K")
    keymap_slack["C-E"] = keymap_slack["F3"]
    keymap_slack["C-K"] = keymap_slack["F3"]
    keymap_slack["F1"] = GENTLE_PUNCHER.invoke("S-SemiColon", "Colon")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")

    def remap_vscode(keys: list, km: WindowKeymap) -> Callable:
        for key in keys:
            km[key] = MILD_PUNCHER.invoke(key)

    remap_vscode(
        [
            "C-E",
            "C-S-F",
            "C-S-E",
            "C-S-G",
            "RC-RS-X",
            "C-0",
            "C-S-P",
            "C-A-B",
        ],
        keymap_vscode,
    )

    # mery
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")

    def remap_mery(mapping_dict: dict, km: WindowKeymap) -> Callable:
        for key, value in mapping_dict.items():
            km[key] = value

    remap_mery(
        {
            "LA-LC-J": "LA-LC-N",
            "LA-LC-K": "LA-LC-LS-N",
            "LA-U0-J": "A-CloseBracket",
            "LA-U0-K": "A-OpenBracket",
            "LA-LC-U0-J": "A-C-CloseBracket",
            "LA-LC-U0-K": "A-C-OpenBracket",
            "LA-LS-U0-J": "A-S-CloseBracket",
            "LA-LS-U0-K": "A-S-OpenBracket",
        },
        keymap_mery,
    )

    # sumatra PDF
    def sumatra_checker(viewmode: bool = False) -> Callable:
        def _checker(wnd: pyauto.Window) -> bool:
            if wnd.getProcessName() == "SumatraPDF.exe":
                if viewmode:
                    return wnd.getClassName() != "Edit"
                return True
            return False

        return _checker

    keymap_sumatra = keymap.defineWindowKeymap(check_func=sumatra_checker(False))

    keymap_sumatra["O-LCtrl"] = "Esc", "Esc", "C-Home", "C-F"

    keymap_sumatra_viewmode = keymap.defineWindowKeymap(check_func=sumatra_checker(True))

    def sumatra_view_key(km: WindowKeymap) -> None:
        for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            km[key] = GENTLE_PUNCHER.invoke(key)

    sumatra_view_key(keymap_sumatra_viewmode)

    keymap_sumatra_viewmode["H"] = "C-S-Tab"
    keymap_sumatra_viewmode["L"] = "C-Tab"

    def office_to_pdf(km: WindowKeymap, key: str = "F11") -> None:
        km[key] = "A-F", "E", "P", "A"

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")
    office_to_pdf(keymap_word)

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    office_to_pdf(keymap_ppt)
    keymap_ppt["O-(236)"] = ImeControl(keymap, 40).to_skk_abbrev

    # excel
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")
    office_to_pdf(keymap_excel)

    def select_all() -> None:
        if keymap.getWindow().getClassName() == "EXCEL6":
            VIRTUAL_FINGER.tap_keys("C-End", "C-S-Home")
        else:
            VIRTUAL_FINGER.tap_keys("C-A")

    keymap_excel["C-A"] = select_all

    def select_cell_content() -> None:
        if keymap.getWindow().getClassName() == "EXCEL7":
            VIRTUAL_FINGER.tap_keys("F2", "C-S-Home")

    keymap_excel["LC-U0-N"] = select_cell_content

    ################################
    # popup clipboard menu
    ################################

    class CharWidth:
        full_letters = "\uff41\uff42\uff43\uff44\uff45\uff46\uff47\uff48\uff49\uff4a\uff4b\uff4c\uff4d\uff4e\uff4f\uff50\uff51\uff52\uff53\uff54\uff55\uff56\uff57\uff58\uff59\uff5a\uff21\uff22\uff23\uff24\uff25\uff26\uff27\uff28\uff29\uff2a\uff2b\uff2c\uff2d\uff2e\uff2f\uff30\uff31\uff32\uff33\uff34\uff35\uff36\uff37\uff38\uff39\uff3a\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19\uff0d"
        half_letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
        full_symbols = "\uff01\uff02\uff03\uff04\uff05\uff06\uff07\uff08\uff09\uff0a\uff0b\uff0c\uff0d\uff0e\uff0f\uff1a\uff1b\uff1c\uff1d\uff1e\uff1f\uff20\uff3b\uff3c\uff3d\uff3e\uff3f\uff40\uff5b\uff5c\uff5d\uff5e"
        half_symbols = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
        full_brackets = "\uff08\uff09\uff3b\uff3d\uff5b\uff5d"
        half_brackets = "()[]{}"

        def __init__(self, totally: bool = False) -> None:
            self._totally = totally

        def to_half_letter(self, s: str) -> str:
            if self._totally:
                return unicodedata.normalize("NFKC", s)
            return s.translate(str.maketrans(self.full_letters, self.half_letters))

        def to_full_letter(self, s: str) -> str:
            s = s.translate(str.maketrans(self.half_letters, self.full_letters))
            if not self._totally:
                return s
            return self.to_full_symbol(s)

        @classmethod
        def to_half_symbol(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.full_symbols, cls.half_symbols))

        @classmethod
        def to_full_symbol(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.half_symbols, cls.full_symbols))

        @classmethod
        def to_half_brackets(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.full_brackets, cls.half_brackets))

        @classmethod
        def to_full_brackets(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.half_brackets, cls.full_brackets))

    def format_zoom_invitation(s: str) -> None:

        def _is_ignorable(line: str) -> bool:
            phrases = [
                "あなたをスケジュール済みの",
                "ミーティングに参加する",
            ]
            for p in phrases:
                if p in line:
                    return True
            return False

        stack = []
        lines = s.strip().splitlines()
        for line in lines:
            if _is_ignorable(line):
                stack.append("")
            else:
                stack.append(line)
        content = "\n".join(stack).strip()
        hr = "=============================="
        return "\n".join([hr, content, hr])

    def md_frontmatter() -> str:
        return "\n".join(
            [
                "---",
                "title:",
                "load:",
                "  - style.css",
                "---",
            ]
        )

    class FormatTools:
        @staticmethod
        def to_deepl_friendly(s: str) -> str:
            ss = []
            lines = s.splitlines()
            for line in lines:
                if line.endswith(" "):
                    ss.append(line)
                elif line.endswith("-"):
                    ss.append(line[0:-1])
                else:
                    ss.append(line + " ")
            return "".join(ss).strip()

        @staticmethod
        def swap_abbreviation(s: str) -> str:
            ss = re.split(r"[:：]\s*", s)
            if len(ss) == 2:
                return ss[1] + "：" + ss[0]
            return ""

        @staticmethod
        def colon_to_doubledash(s: str) -> str:
            return re.sub(r"[:：]\s*", "\u2015\u2015", s)

        @staticmethod
        def as_codeblock(s: str) -> str:
            return "\n".join(["```", s, "```", ""])

        @staticmethod
        def skip_blank_line(s: str) -> str:
            lines = s.strip().splitlines()
            return "\n".join([l for l in lines if l.strip()])

        @staticmethod
        def insert_blank_line(s: str) -> str:
            lines = []
            for line in s.strip().splitlines():
                lines.append(line.strip())
                lines.append("")
            return "\n".join(lines)

        @staticmethod
        def to_double_bracket(s: str) -> str:
            reg = re.compile(r"[\u300c\u300d]")

            def _replacer(mo: re.Match) -> str:
                if mo.group(0) == "\u300c":
                    return "\u300e"
                return "\u300f"

            return reg.sub(_replacer, s)

        @staticmethod
        def to_single_bracket(s: str) -> str:
            reg = re.compile(r"[\u300e\u300f]")

            def _replacer(mo: re.Match) -> str:
                if mo.group(0) == "\u300e":
                    return "\u300c"
                return "\u300d"

            return reg.sub(_replacer, s)

        @staticmethod
        def split_postalcode(s: str) -> str:
            lines = s.splitlines()
            if 1 < len(lines):
                reg = re.compile(r"(\d{3}).(\d{4})[ 　]*(.+$)")
            else:
                reg = re.compile(r"(\d{3}).(\d{4})[\s]*(.+$)")
            ss = []
            for line in lines:
                hankaku = CharWidth().to_half_letter(line.strip().strip("\u3012"))
                m = reg.match(hankaku)
                if m:
                    ss.append("{}-{}\t{}".format(m.group(1), m.group(2), m.group(3)))
                else:
                    ss.append(line)
            return "\n".join(ss)

        @staticmethod
        def fix_paren_inside_bracket(s: str) -> str:
            reg = re.compile(r"(\(.+?\)|（.+?）)」")

            def _replacer(mo: re.Match) -> str:
                return "」" + mo.group(1)

            return reg.sub(_replacer, s)

        @staticmethod
        def fix_dumb_quotation(s: str) -> str:
            reg = re.compile(r"\"([^\"]+?)\"|'([^']+?)'")

            def _replacer(mo: re.Match) -> str:
                if str(mo.group(0)).startswith('"'):
                    return "\u201c{}\u201d".format(mo.group(1))
                return "\u2018{}\u2019".format(mo.group(1))

            return reg.sub(_replacer, s)

        @staticmethod
        def decode_url(s: str) -> str:
            return urllib.parse.unquote(s)

        @staticmethod
        def encode_url(s: str) -> str:
            return urllib.parse.quote(s)

        @staticmethod
        def trim_honorific(s: str) -> str:
            reg = re.compile(r"先生$|様$|(先生|様)(?=[、。：；（）［］・！？])")
            return reg.sub("", s)

        @staticmethod
        def mdtable_from_tsv(s: str) -> str:
            delim = "\t"

            def _split(s: str) -> str:
                return s.split(delim)

            def _join(ss: list) -> str:
                pipe = "|"
                return pipe + pipe.join(ss) + pipe

            lines = s.splitlines()
            header = _join(_split(lines[0]))
            sep = _join([":---:" for _ in lines[0].split(delim)])
            table = [
                header,
                sep,
            ]
            for line in lines[1:]:
                table.append(_join(_split(line)))
            return "\n".join(table)

    class ClipboardMenu(ClipHandler):
        def __init__(self) -> None:
            self._table = {}

        @classmethod
        def invoke_formatter(cls, func: Callable) -> Callable:
            def _formatter() -> str:
                cb = cls.get_string()
                if cb:
                    return func(cb)

            return _formatter

        @classmethod
        def invoke_replacer(cls, search: str, replace_to: str) -> Callable:
            reg = re.compile(search)

            def _replacer() -> str:
                cb = cls.get_string()
                if cb:
                    return reg.sub(replace_to, cb)

            return _replacer

        @property
        def table(self) -> dict:
            return self._table

        def set_formatter(self, mapping: dict) -> None:
            for menu, func in mapping.items():
                self._table[menu] = self.invoke_formatter(func)

        def set_replacer(self, mapping: dict) -> None:
            for menu, args in mapping.items():
                self._table[menu] = self.invoke_replacer(*args)

        def set_func(self, mapping: dict) -> None:
            for menu, func in mapping.items():
                self._table[menu] = func

    CLIPBOARD_MENU = ClipboardMenu()
    CLIPBOARD_MENU.set_formatter(
        {
            "to deepl-friendly": FormatTools.to_deepl_friendly,
            "swap abbreviation around colon": FormatTools.swap_abbreviation,
            "colon to double-dash": FormatTools.colon_to_doubledash,
            "insert blank line": FormatTools.insert_blank_line,
            "remove blank line": FormatTools.skip_blank_line,
            "fix dumb quotation": FormatTools.fix_dumb_quotation,
            "fix KANGXI RADICALS": KangxiRadicals().fix,
            "fix paren inside bracket": FormatTools.fix_paren_inside_bracket,
            "to double bracket": FormatTools.to_double_bracket,
            "to single bracket": FormatTools.to_single_bracket,
            "to markdown codeblock": FormatTools.as_codeblock,
            "TSV to markdown table": FormatTools.mdtable_from_tsv,
            "split postalcode and address": FormatTools.split_postalcode,
            "decode url": FormatTools.decode_url,
            "encode url": FormatTools.encode_url,
            "to halfwidth": CharWidth().to_half_letter,
            "to halfwidth (including symbols)": CharWidth(True).to_half_letter,
            "to halfwidth symbols": CharWidth().to_half_symbol,
            "to halfwidth bracktets": CharWidth().to_half_brackets,
            "to fullwidth": CharWidth().to_full_letter,
            "to fullwidth (including symbols)": CharWidth(True).to_full_letter,
            "to fullwidth symbols": CharWidth().to_full_symbol,
            "to fullwidth bracktets": CharWidth().to_full_brackets,
            "trim honorific": FormatTools.trim_honorific,
            "zoom invitation": format_zoom_invitation,
        }
    )
    CLIPBOARD_MENU.set_replacer(
        {
            "escape backslash": (r"\\", r"\\\\"),
            "escape double-quotation": (r"\"", r'\\"'),
            "remove double-quotation": (r'"', ""),
            "remove single-quotation": (r"'", ""),
            "remove linebreak": (r"\r?\n", ""),
            "to sigle line": (r"\r?\n", ""),
            "remove whitespaces": (
                r"[\u0009\u0020\u00a0\u2000-\u200f\u202f\u205f\u3000\ufeff]",
                "",
            ),
            "remove whitespaces (including linebreak)": (r"\s", ""),
            "remove non-digit-char": (r"[^\d]", ""),
            "remove quotations": (r"[\u0022\u0027]", ""),
            "remove inside paren": (r"[（\(].+?[）\)]", ""),
            "fix msword-bullet": (r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009", "\u30fb"),
            "remove msword-bullet": (r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009", ""),
            "to curly-comma (\uff0c)": (r"\u3001", "\uff0c"),
            "to japanese-comma (\u3001)": (r"\uff0c", "\u3001"),
            "shorten amazon url": (
                r"^.+amazon\.co\.jp/.+dp/(.{10}).*",
                r"https://www.amazon.jp/dp/\1",
            ),
        }
    )
    CLIPBOARD_MENU.set_func(
        {
            "to lowercase": lambda: ClipHandler.get_string().lower(),
            "to uppercase": lambda: ClipHandler.get_string().upper(),
            "my markdown frontmatter": md_frontmatter,
        }
    )

    def fzfmenu() -> None:
        if not check_fzf():
            balloon("cannot find fzf on PC.")
            return

        if not ClipHandler.get_string():
            balloon("no text in clipboard.")
            return

        table = CLIPBOARD_MENU.table

        def _fzf(job_item: ckit.JobItem) -> None:
            job_item.result = False
            job_item.paste_string = ""
            job_item.skip_paste = False
            proc = subprocess.Popen(
                ["fzf.exe"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
            )
            try:
                for k in table.keys():
                    proc.stdin.write(k + "\n")
                proc.stdin.close()
            except Exception as e:
                balloon(e)
                return

            result, err = proc.communicate()
            if proc.returncode != 0:
                if err:
                    print(err)
                return
            result = result.strip()
            if len(result) < 1:
                return

            func = table.get(result, None)
            if func:
                fmt = func()
                if 0 < len(fmt):
                    job_item.result = True
                    job_item.paste_string = fmt

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.result and job_item.paste_string:
                ClipHandler.paste(job_item.paste_string)

        subthread_run(_fzf, _finished)

    keymap_global["U1-Z"] = LAZY_KEYMAP.defer(fzfmenu, 120)


def configure_ListWindow(window: ListWindow) -> None:
    window.keymap["J"] = window.command_CursorDown
    window.keymap["K"] = window.command_CursorUp
    window.keymap["C-J"] = window.command_CursorPageDown
    window.keymap["C-K"] = window.command_CursorPageUp
    window.keymap["L"] = window.command_Enter
    for mod in ["", "S-", "C-", "C-S-"]:
        for key in ["L", "Space"]:
            window.keymap[mod + key] = window.command_Enter
