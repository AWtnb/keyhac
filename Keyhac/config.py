import datetime
import os
import sys
import fnmatch
import re
import time
import subprocess
import urllib.parse
import unicodedata
from typing import Union, Callable, Dict, List, NamedTuple
from pathlib import Path
from winreg import HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, OpenKey, QueryValueEx

import ckit
import pyauto
from keyhac import *
from keyhac_keymap import Keymap, KeyCondition, WindowKeymap, VK_CAPITAL
from keyhac_listwindow import ListWindow


def smart_check_path(path: Union[str, Path]) -> bool:
    # case insensitive path check
    p = Path(path) if type(path) is str else path
    if p.drive == "C:":
        return p.exists()
    return os.path.exists(path)


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
            self._path = path

        @property
        def path(self) -> str:
            return self._path

        def is_accessible(self) -> bool:
            if self._path:
                try:
                    return smart_check_path(self._path)
                except Exception as e:
                    print(e)
                    return ""
            return False

        @staticmethod
        def args_to_param(args: tuple) -> str:
            params = []
            for arg in args:
                if len(arg.strip()):
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

    class UserPath(PathHandler):
        def __init__(self, rel: str) -> None:
            path = self.resolve(rel)
            super().__init__(path)

        @staticmethod
        def resolve(rel: str) -> str:
            user_prof = os.environ.get("USERPROFILE") or ""
            return str(Path(user_prof, rel))

    def get_editor() -> str:
        vscode_path = UserPath(r"scoop\apps\vscode\current\Code.exe")
        if vscode_path.is_accessible():
            return vscode_path.path
        return "notepad.exe"

    KEYHAC_EDITOR = get_editor()

    # console theme
    keymap.setFont("HackGen", 16)

    KEYHAC_THEME = "black"

    class Themer:
        def __init__(self, color: str) -> None:

            if color == "black":
                colortable = {
                    "bg": (0, 0, 0),
                    "fg": (255, 255, 255),
                    "cursor0": (255, 255, 255),
                    "cursor1": (255, 64, 64),
                    "bar_fg": (0, 0, 0),
                    "bar_error_fg": (200, 0, 0),
                    "select_bg": (30, 100, 150),
                    "select_fg": (255, 255, 255),
                }
            else:
                colortable = {
                    "bg": (255, 255, 255),
                    "fg": (0, 0, 0),
                    "cursor0": (255, 255, 255),
                    "cursor1": (0, 255, 255),
                    "bar_fg": (255, 255, 255),
                    "bar_error_fg": (255, 0, 0),
                    "select_bg": (70, 200, 255),
                    "select_fg": (0, 0, 0),
                }
            self._theme_path = Path(ckit.getAppExePath(), "theme", color, "theme.ini")
            self._data = colortable

        def update(self, key: str, value: str) -> None:
            if key not in self._data:
                return
            if type(value) is tuple:
                self._data[key] = value
                return
            colorcode = value.strip("#")
            if len(colorcode) == 6:
                r, g, b = colorcode[:2], colorcode[2:4], colorcode[4:6]
                try:
                    rgb = tuple(int(c, 16) for c in [r, g, b])
                    self._data[key] = rgb
                except Exception as e:
                    print(e)

        def to_string(self) -> str:
            lines = ["[COLOR]"]
            for key, value in self._data.items():
                line = "{} = {}".format(key, value)
                lines.append(line)
            return "\n".join(lines)

        def overwrite(self) -> None:
            theme = self.to_string()
            if not smart_check_path(self._theme_path) or self._theme_path.read_text() != theme:
                self._theme_path.write_text(theme)

    def set_theme(theme_table: dict):
        t = Themer(KEYHAC_THEME)
        for k, v in theme_table.items():
            t.update(k, v)
        t.overwrite()
        keymap.setTheme(KEYHAC_THEME)

    CUSTOM_THEME = {
        "bg": "#3F3B39",
        "fg": "#A0B4A7",
        "cursor0": "#FFFFFF",
        "cursor1": "#FF4040",
        "bar_fg": "#000000",
        "bar_error_fg": "#FF4040",
        "select_bg": "#DFF477",
        "select_fg": "#3F3B39",
    }
    set_theme(CUSTOM_THEME)

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

    # keymap working on any window
    keymap_global = keymap.defineWindowKeymap(check_func=CheckWnd.is_global_target)

    # clipboard menu
    keymap_global["LC-LS-X"] = keymap.command_ClipboardList

    # keyboard macro
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["S-U0-0"] = keymap.command_RecordClear
    keymap_global["U1-0"] = keymap.command_RecordPlay
    keymap_global["U1-F4"] = keymap.command_RecordPlay
    keymap_global["C-U0-0"] = keymap.command_RecordPlay

    # combination with modifier key
    class CoreKeys:
        mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
        key_status = ("D-", "U-")

        @classmethod
        def cursor_keys(cls, km: WindowKeymap) -> None:
            for mod_key in cls.mod_keys:
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
                    km[mod_key + "U0-" + key] = mod_key + value

        @classmethod
        def ignore_capslock(cls, km: WindowKeymap) -> None:
            for stat in cls.key_status:
                for mod_key in cls.mod_keys:
                    km[mod_key + stat + "Capslock"] = lambda: None

        @classmethod
        def ignore_kanakey(cls, km: WindowKeymap) -> None:
            for stat in cls.key_status:
                for mod_key in cls.mod_keys:
                    for vk in list(range(124, 136)) + list(range(240, 243)) + list(range(245, 254)):
                        km[mod_key + stat + str(vk)] = lambda: None

    CoreKeys().cursor_keys(keymap_global)
    # CoreKeys().ignore_capslock(keymap_global)
    # CoreKeys().ignore_kanakey(keymap_global)

    class KeyAllocator:
        def __init__(self, mapping_dict: dict) -> None:
            self.dict = mapping_dict

        def apply(self, km: WindowKeymap):
            for key, value in self.dict.items():
                km[key] = value

        def apply_quotation(self, km: WindowKeymap):
            for key, value in self.dict.items():
                km[key] = value, value, "Left"

    KeyAllocator(
        {
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
            "U1-E": ("S-Minus"),
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
    ).apply(keymap_global)

    KeyAllocator(
        {
            "U0-2": "LS-2",
            "U0-7": "LS-7",
            "U0-AtMark": "LS-AtMark",
        }
    ).apply_quotation(keymap_global)

    ################################
    # functions for custom hotkey
    ################################

    def delay(msec: int = 50) -> None:
        time.sleep(msec / 1000)

    class Key(NamedTuple):
        sent: str
        typable: bool

    class KeySequence:
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
        def wrap(cls, sequence: tuple) -> List[Key]:
            seq = []
            for elem in sequence:
                key = Key(elem, cls.check(elem))
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

        def type_keys(self, *keys) -> None:
            self._prepare()
            for key in keys:
                delay(self._inter_stroke_pause)
                self._keymap.setInput_FromString(str(key))
            self._finish()

        def type_text(self, s: str) -> None:
            self._prepare()
            for c in str(s):
                delay(self._inter_stroke_pause)
                self._keymap.input_seq.append(pyauto.Char(c))
            self._finish()

        def type_smart(self, *sequence) -> None:
            for elem in sequence:
                try:
                    self.type_keys(elem)
                except:
                    self.type_text(elem)

        def type_sequence(self, sequence: List[Key]) -> None:
            for key in sequence:
                if key.typable:
                    self.type_keys(key.sent)
                else:
                    self.type_text(key.sent)

    VIRTUAL_FINGER = VirtualFinger(keymap, 10)
    VIRTUAL_FINGER_QUICK = VirtualFinger(keymap, 0)

    class ImeControl:
        kata_key = "Q"
        kana_key = "C-J"
        latin_key = "S-L"
        cancel_key = "Esc"
        reconv_key = "LWin-Slash"
        abbrev_key = "Slash"

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
            self._finger.type_keys(self.kana_key)

        def to_skk_latin(self) -> None:
            self.enable_skk()
            self._finger.type_keys(self.latin_key)

        def to_skk_abbrev(self) -> None:
            self.enable_skk()
            self._finger.type_keys(self.abbrev_key)

        def to_skk_kata(self) -> None:
            self.enable_skk()
            self._finger.type_keys(self.kata_key)

        def reconvert_with_skk(self) -> None:
            self.enable_skk()
            self._finger.type_keys(self.reconv_key, self.cancel_key)

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
            VIRTUAL_FINGER.type_keys("C-V")

        @classmethod
        def paste_current(cls, format_func: Union[Callable, None] = None) -> None:
            cls.paste(cls.get_string(), format_func)

        @classmethod
        def after_copy(cls, deferred: Callable) -> None:
            cb = cls.get_string()
            VIRTUAL_FINGER.type_keys("C-C")

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
                cls.set_string(job_item.origin + os.linesep + job_item.copied)

            cls.after_copy(_push)

    class LazyFunc:
        def __init__(self, keymap: Keymap, func: Callable) -> None:
            self._keymap = keymap

            def _wrapper() -> None:
                self._keymap.hookCall(func)

            self._func = _wrapper

        def defer(self, msec: int = 20) -> Callable:
            def _executer() -> None:
                # use delayedCall in ckit => https://github.com/crftwr/ckit/blob/2ea84f986f9b46a3e7f7b20a457d979ec9be2d33/ckitcore/ckitcore.cpp#L1998
                self._keymap.delayedCall(self._func, msec)

            return _executer

    class LazyKeymap:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        def wrap(self, func: Callable) -> LazyFunc:
            return LazyFunc(self._keymap, func)

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
            self._finger = VirtualFinger(keymap, inter_stroke_pause)
            self._control = ImeControl(keymap)
            self._lazy_keymap = LazyKeymap(keymap)

        def invoke(self, *sequence) -> Callable:
            seq = KeySequence().wrap(sequence)

            def _input() -> None:
                self._control.disable()
                self._finger.type_sequence(seq)
                if self._recover_ime:
                    self._control.enable()

            return self._lazy_keymap.wrap(_input).defer(self._defer_msec)

    MILD_PUNCHER = KeyPuncher(keymap, defer_msec=20)
    GENTLE_PUNCHER = KeyPuncher(keymap, defer_msec=50)

    keymap_global["LC-Q"] = GENTLE_PUNCHER.invoke("A-F4")
    keymap_global["U0-4"] = GENTLE_PUNCHER.invoke("$_")
    keymap_global["U1-4"] = GENTLE_PUNCHER.invoke("$_.")

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
    keymap_global["U0-(236)"] = IME_CONTROL.enable_skk
    keymap_global["U1-(235)"] = IME_CONTROL.disable

    # paste as plaintext
    keymap_global["U0-V"] = LAZY_KEYMAP.wrap(ClipHandler().paste_current).defer()

    # paste as plaintext (with trimming removable whitespaces)
    class StrCleaner:
        @staticmethod
        def clear_space(s: str) -> str:
            return s.strip().translate(str.maketrans("", "", "\u200b\u3000\u0009\u0020\u00a0"))

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
            return os.linesep.join(["> " + line for line in lines])

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
            self._keymap.console_window.reloadTheme()
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
                km[key] = LAZY_KEYMAP.wrap(func).defer(50)

    CONFIG_MENU = ConfigMenu(keymap)
    keymap_global["LC-U0-X"] = keymap.defineMultiStrokeKeymap()
    CONFIG_MENU.apply(keymap_global["LC-U0-X"])

    keymap.editor = lambda _: CONFIG_MENU.open_keyhac_repo()

    keymap_global["U1-F12"] = LAZY_KEYMAP.wrap(ConfigMenu(keymap).reload_config).defer(50)

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

        def __init__(self, keymap: keymap) -> None:
            self._keymap = keymap

        def alloc_flexible(self, km: WindowKeymap) -> None:
            monitors = CurrentMonitors(self._keymap).monitors
            for mod_mntr, mntr_idx in self.monitor_dict.items():
                for mod_area, size in self.size_dict.items():
                    for key, pos in self.snap_key_dict.items():
                        if mntr_idx < len(monitors):
                            wnd_rect = monitors[mntr_idx].area_mapping[pos][size]
                            km[mod_mntr + mod_area + key] = LAZY_KEYMAP.wrap(wnd_rect.snap).defer(
                                50
                            )

        def alloc_maximize(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, towards in mapping_dict.items():

                def _snap() -> None:
                    def _maximize(_) -> None:
                        self._keymap.getTopLevelWindow().maximize()

                    def _snapper(_) -> None:
                        VIRTUAL_FINGER.type_keys("LShift-LWin-" + towards)

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

            return LAZY_KEYMAP.wrap(_snapper).defer(50)

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

    class SimpleSKK:
        def __init__(
            self,
            keymap: Keymap,
            inter_stroke_pause: int = 0,
        ) -> None:
            self._finger = VirtualFinger(keymap, inter_stroke_pause)
            self._control = ImeControl(keymap)

        def under_kanamode(self, *sequence) -> Callable:
            seq = KeySequence().wrap(sequence)

            def _send() -> None:
                self._control.enable_skk()
                self._finger.type_sequence(seq)

            return _send

        def under_latinmode(self, *sequence) -> Callable:
            seq = KeySequence().wrap(sequence)

            def _send() -> None:
                self._control.to_skk_latin()
                self._finger.type_sequence(seq)

            return _send

        def disable(self) -> None:
            self._control.disable()

    SIMPLE_SKK = SimpleSKK(keymap)

    # select-to-left with ime control
    keymap_global["U1-B"] = SIMPLE_SKK.under_kanamode("S-Left")
    keymap_global["LS-U1-B"] = SIMPLE_SKK.under_kanamode("S-Right")
    keymap_global["U1-Space"] = SIMPLE_SKK.under_kanamode("C-S-Left")
    keymap_global["U1-N"] = SIMPLE_SKK.under_kanamode("S-Left", ImeControl.abbrev_key)
    keymap_global["U1-Tab"] = SIMPLE_SKK.under_kanamode("End", "S-Home")
    keymap_global["U0-Tab"] = SIMPLE_SKK.under_latinmode("End", "S-Home")

    class SKKMode:
        disabled = -1
        latin = 0
        kana = 1

    class SKK:
        def __init__(self, keymap: Keymap, finish_mode: SKKMode = SKKMode.kana) -> None:
            self._base_skk = SimpleSKK(keymap)
            self._finish_mode = finish_mode

        def invoke_sender(self, *sequence) -> Callable:
            if self._finish_mode == SKKMode.kana:
                sequence = list(sequence) + [ImeControl.kana_key]

            sender = self._base_skk.under_latinmode(*sequence)
            if self._finish_mode == SKKMode.disabled:

                def _sender():
                    sender()
                    self._base_skk.disable()

                return _sender
            return sender

        def invoke_pair_sender(self, pair: list) -> Callable:
            _, suffix = pair
            sequence = pair + ["Left"] * len(suffix)
            return self.invoke_sender(*sequence)

        def apply(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, sent in mapping_dict.items():
                km[key] = self.invoke_sender(sent)

        def apply_pair(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, sent in mapping_dict.items():
                km[key] = self.invoke_pair_sender(sent)

    SKK_TO_KANAMODE = SKK(keymap, SKKMode.kana)
    SKK_TO_LATINMODE = SKK(keymap, SKKMode.latin)
    SKK_TO_DISABLE = SKK(keymap, SKKMode.disabled)

    # insert honorific
    def type_honorific(km: WindowKeymap) -> None:
        for key, hono in {"U0": "先生", "U1": "様"}.items():
            for mod, suffix in {"": "", "C-": "方"}.items():
                km[mod + key + "-Tab"] = SKK_TO_KANAMODE.invoke_sender(hono + suffix)

    type_honorific(keymap_global)

    # markdown list
    keymap_global["S-U0-8"] = SKK_TO_KANAMODE.invoke_sender("- ")
    keymap_global["U1-1"] = SKK_TO_KANAMODE.invoke_sender("1. ")

    SKK_TO_KANAMODE.apply(
        keymap_global,
        {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Minus": "\u3000\u2015\u2015",
            "S-U0-Period": "\uff0e",  # FULLWIDTH FULL STOP
            "U0-Minus": "\u2015\u2015",  # HORIZONTAL BAR * 2
            "U0-P": "\u30fb",  # KATAKANA MIDDLE DOT
            "S-U0-SemiColon": "+ ",
        },
    )
    SKK_TO_KANAMODE.apply_pair(
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

    SKK_TO_LATINMODE.apply(
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
    SKK_TO_LATINMODE.apply_pair(
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

    class DateInput:
        def __init__(self) -> None:
            pass

        @staticmethod
        def invoke(fmt: str, finish_with_kanamode: bool = False) -> Callable:
            def _inputter() -> None:
                def _get_func(job_item: ckit.JobItem) -> None:
                    d = datetime.datetime.today()
                    seq = [c for c in d.strftime(fmt)]
                    if finish_with_kanamode:
                        job_item.func = SKK_TO_KANAMODE.invoke_sender(*seq)
                    else:
                        job_item.func = SKK_TO_LATINMODE.invoke_sender(*seq)

                def _input(job_item: ckit.JobItem) -> None:
                    job_item.func()

                subthread_run(_get_func, _input)

            return _inputter

        def apply(self, km: WindowKeymap) -> None:
            for key, params in {
                "1": ("%Y%m%d", False),
                "2": ("%Y/%m/%d", False),
                "3": ("%Y.%m.%d", False),
                "4": ("%Y-%m-%d", False),
                "5": ("%Y年%#m月%#d日", True),
                "D": ("%Y%m%d", False),
                "S": ("%Y/%m/%d", False),
                "P": ("%Y.%m.%d", False),
                "H": ("%Y-%m-%d", False),
                "U": ("%Y_%m_%d", False),
                "M": ("%Y%m", False),
                "J": ("%Y年%#m月%#d日", True),
            }.items():
                km[key] = self.invoke(*params)

    keymap_global["U1-D"] = keymap.defineMultiStrokeKeymap()
    DateInput().apply(keymap_global["U1-D"])

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

            return LAZY_KEYMAP.wrap(_searcher).defer()

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
            "Y": "http://www.google.co.jp/search?q=site%3Ayuhikaku.co.jp%20{}",
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

        def scan(self) -> None:
            self.found = None
            pyauto.Window.enum(self.traverse_wnd, None)

        def traverse_wnd(self, wnd: pyauto.Window, _) -> bool:
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
            self.found = wnd.getLastActivePopup()
            return False

    class PseudoCuteExec:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap

        def activate_wnd(self, target: pyauto.Window) -> bool:
            interval = 20
            timeout = interval * 50
            while timeout > 0:
                try:
                    target.setForeground()
                    delay(interval)
                    if pyauto.Window.getForeground() == target:
                        if target.isMinimized():
                            target.restore()
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
                        VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")

                subthread_run(_activate, _finished)

            return _executer

        def apply(self, wnd_keymap: WindowKeymap, remap_table: dict = {}) -> None:
            for key, params in remap_table.items():
                func = self.invoke(*params)
                wnd_keymap[key] = LAZY_KEYMAP.wrap(func).defer(10)

    PSEUDO_CUTEEXEC = PseudoCuteExec(keymap)

    PSEUDO_CUTEEXEC.apply(
        keymap_global,
        {
            "U1-F": (
                "cfiler.exe",
                "CfilerWindowClass",
                UserPath.resolve(r"Sync\portable_app\cfiler\cfiler.exe"),
            ),
            "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
            "LC-U1-M": (
                "Mery.exe",
                "TChildForm",
                UserPath.resolve(r"AppData\Local\Programs\Mery\Mery.exe"),
            ),
            "LC-U1-N": (
                "notepad.exe",
                "Notepad",
                r"C:\Windows\System32\notepad.exe",
            ),
            "LC-AtMark": (
                "wezterm-gui.exe",
                "org.wezfurlong.wezterm",
                UserPath.resolve(r"scoop\apps\wezterm\current\wezterm-gui.exe"),
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
                UserPath.resolve(r"AppData\Local\Vivaldi\Application\vivaldi.exe"),
            ),
            "S": (
                "slack.exe",
                "Chrome_WidgetWin_1",
                UserPath.resolve(r"AppData\Local\slack\slack.exe"),
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
                UserPath.resolve(r"scoop\apps\ksnip\current\ksnip.exe"),
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
                UserPath.resolve(r"AppData\Local\Programs\Mery\Mery.exe"),
            ),
            "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
        },
    )

    keymap_global["LS-LC-U1-M"] = UserPath(r"Personal\draft.txt").run

    def search_on_browser() -> None:
        if keymap.getWindow().getProcessName() == DEFAULT_BROWSER.get_exe_name():
            VIRTUAL_FINGER.type_keys("C-T")
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
                VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")
                return
            VIRTUAL_FINGER.type_keys("C-T")

        subthread_run(_activate, _finished)

    keymap_global["U0-Q"] = LAZY_KEYMAP.wrap(search_on_browser).defer(10)

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

    keymap_sumatra_viewmode["F"] = SIMPLE_SKK.under_kanamode("C-F")
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
            VIRTUAL_FINGER.type_keys("C-End", "C-S-Home")
        else:
            VIRTUAL_FINGER.type_keys("C-A")

    keymap_excel["C-A"] = select_all

    def select_cell_content() -> None:
        if keymap.getWindow().getClassName() == "EXCEL7":
            VIRTUAL_FINGER.type_keys("F2", "C-S-Home")

    keymap_excel["LC-U0-N"] = select_cell_content

    # Thunderbird
    def thunderbird_new_mail(sequence: list, alt_sequence: list) -> Callable:
        def _sender():
            wnd = keymap.getWindow()
            if wnd.getProcessName() == "thunderbird.exe":
                if wnd.getText().startswith("作成: (件名なし)"):
                    VIRTUAL_FINGER.type_keys(*sequence)
                else:
                    VIRTUAL_FINGER.type_keys(*alt_sequence)

        return _sender

    keymap_tb = keymap.defineWindowKeymap(exe_name="thunderbird.exe")
    keymap_tb["C-S-V"] = thunderbird_new_mail(["A-S", "Tab", "C-V", "C-Home"], ["C-V"])
    keymap_tb["C-S-S"] = thunderbird_new_mail(
        ["C-Home", "S-End", "C-X", "Delete", "A-S", "C-V"], ["A-S"]
    )

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

    class Zoom:
        separator = ": "
        hr = "=============================="

        @staticmethod
        def get_time(s) -> str:
            try:
                d = datetime.datetime.strptime(s, "時刻: %Y年%m月%d日 %I:%M %p 大阪、札幌、東京")
            except:
                try:
                    d = datetime.datetime.strptime(s, "時刻: %Y年%m月%d日 %H:%M 大阪、札幌、東京")
                except:
                    return ""
            week = "月火水木金土日"[d.weekday()]
            ampm = ""
            if d.hour < 12:
                ampm = "AM "
            return (d.strftime("%Y年%m月%d日（{}） {}%H:%M開始")).format(week, ampm)

        @classmethod
        def to_field(cls, s: str, prefix: str) -> str:
            i = s.find(cls.separator)
            if -1 < i:
                v = s[(i + len(cls.separator)) :].strip()
            else:
                v = s
            if 0 < len(prefix):
                c = ": "
            else:
                c = ""
            return prefix + c + v

        @classmethod
        def format(cls, copied: str) -> str:
            lines = copied.strip().splitlines()
            if len(lines) < 9:
                print("Zoom format ERROR: lack of lines.")
                return copied
            due = cls.get_time(lines[3])
            if len(due) < 1:
                print("Zoom format ERROR: could not parse due date.")
                return copied
            return os.linesep.join(
                [
                    cls.hr,
                    cls.to_field(lines[2], ""),
                    cls.to_field(due, ""),
                    cls.to_field(lines[6], ""),
                    "",
                    cls.to_field(lines[8], "meeting ID"),
                    cls.to_field(lines[9], "passcode"),
                    cls.hr,
                ]
            )

    def md_frontmatter() -> str:
        return os.linesep.join(
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
            return os.linesep.join([l for l in lines if l.strip()])

        @staticmethod
        def insert_blank_line(s: str) -> str:
            lines = []
            for line in s.strip().splitlines():
                lines.append(line.strip())
                lines.append("")
            return os.linesep.join(lines)

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
            return os.linesep.join(ss)

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
            return os.linesep.join(table)

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
            "zoom invitation": Zoom().format,
        }
    )
    CLIPBOARD_MENU.set_replacer(
        {
            "escape backslash": (r"\\", r"\\\\"),
            "escape double-quotation": (r"\"", r'\\"'),
            "remove double-quotation": (r'"', ""),
            "remove single-quotation": (r"'", ""),
            "remove linebreak": (r"\r?\n", ""),
            "remove whitespaces": (r"[\u200b\u3000\u0009\u0020]", ""),
            "remove whitespaces (including linebreak)": (r"\s", ""),
            "remove non-digit-char": (r"[^\d]", ""),
            "remove quotations": (r"[\u0022\u0027]", ""),
            "remove inside paren": (r"[（\(].+?[）\)]", ""),
            "fix msword-bullet": (r"\uf09f\u0009", "\u30fb"),
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

    def check_fzf() -> bool:
        paths = os.environ.get("PATH", "").split(os.pathsep)
        for path in paths:
            if Path(path, "fzf.exe").exists():
                return True
        return False

    class FzfResult(NamedTuple):
        finisher: str
        text: str

    class FzfResultParser:
        default_finisher = "enter"

        def __init__(self, expect: bool) -> None:
            self.expect = expect

        def parse(self, stdout: str) -> FzfResult:
            if len(stdout) < 1:
                return FzfResult(self.default_finisher, "")
            if not self.expect:
                return FzfResult(self.default_finisher, stdout)
            lines = stdout.splitlines()
            try:
                assert (
                    len(lines) == 2
                ), "with --expect option, 2 lines should be returned, but 3 or more lines are returned"
            except AssertionError as err:
                print(err, file=sys.stderr)
                return FzfResult(self.default_finisher, "")
            if len(lines[0]) < 1:
                return FzfResult(self.default_finisher, lines[1])
            return FzfResult(*lines)

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
            lines = "\n".join(table.keys())
            proc = subprocess.run(
                ["fzf.exe", "--expect", "ctrl-space"],
                input=lines,
                capture_output=True,
                encoding="utf-8",
            )
            result = proc.stdout
            if len(result) < 1 or proc.returncode != 0:
                return
            fr = FzfResultParser(True).parse(result)
            result_func = fr.text
            skip = fr.finisher != FzfResultParser.default_finisher
            func = table.get(result_func, None)
            if func:
                fmt = func()
                if 0 < len(fmt):
                    job_item.result = True
                    job_item.paste_string = fmt
                    job_item.skip_paste = skip

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.result and job_item.paste_string:
                if job_item.skip_paste:
                    ClipHandler.set_string(job_item.paste_string)
                else:
                    ClipHandler.paste(job_item.paste_string)

        subthread_run(_fzf, _finished)

    keymap_global["U1-Z"] = LAZY_KEYMAP.wrap(fzfmenu).defer(80)


def configure_ListWindow(window: ListWindow) -> None:
    window.keymap["J"] = window.command_CursorDown
    window.keymap["K"] = window.command_CursorUp
    window.keymap["C-J"] = window.command_CursorPageDown
    window.keymap["C-K"] = window.command_CursorPageUp
    window.keymap["L"] = window.command_Enter
    for mod in ["", "S-", "C-", "C-S-"]:
        for key in ["L", "Space"]:
            window.keymap[mod + key] = window.command_Enter
