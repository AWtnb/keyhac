import datetime
import os
import fnmatch
import re
import time
import urllib.parse
from typing import Union, Callable
from pathlib import Path
from winreg import HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, OpenKey, QueryValueEx

import pyauto
from keyhac import *
from keyhac_keymap import Keymap, WindowKeymap


def configure(keymap):
    ################################
    # general setting
    ################################

    class PathHandler:
        def __init__(self, s: str, under_user_profile: bool = False) -> None:
            if under_user_profile:
                self._path = self.resolve_user_profile(s)
            else:
                self._path = s

        @property
        def path(self) -> str:
            return self._path

        def is_accessible(self) -> bool:
            if self._path:
                return self._path.startswith("http") or Path(self._path).exists()
            return False

        @staticmethod
        def resolve_user_profile(rel) -> str:
            user_prof = os.environ.get("USERPROFILE") or ""
            return str(Path(user_prof, rel))

        @staticmethod
        def args_to_param(args: tuple) -> str:
            params = []
            for arg in args:
                if len(arg.strip()):
                    if " " in arg:
                        params.append('"{}"'.format(arg))
                    else:
                        params.append("{}".format(arg))
            return " ".join(params)

        def run(self, *args) -> None:
            if self.is_accessible():
                keymap.ShellExecuteCommand(None, self._path, self.args_to_param(args), None)()
            else:
                print("invalid-path: '{}'".format(self._path))

    class KeyhacEnv:
        @staticmethod
        def get_filer() -> str:
            tablacus = PathHandler(r"Sync\portable_app\tablacus\TE64.exe", True)
            if tablacus.is_accessible():
                return tablacus.path
            return "explorer.exe"

        @staticmethod
        def get_editor() -> str:
            vscode_path = PathHandler(r"scoop\apps\vscode\current\Code.exe", True)
            if vscode_path.is_accessible():
                return vscode_path.path
            return "notepad.exe"

    KEYHAC_FILER = KeyhacEnv.get_filer()
    KEYHAC_EDITOR = KeyhacEnv.get_editor()

    keymap.editor = KEYHAC_EDITOR

    # console theme
    keymap.setFont("HackGen", 16)
    keymap.setTheme("black")

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
        def is_filer_viewmode(wnd: pyauto.Window) -> bool:
            if wnd.getProcessName() == "explorer.exe":
                return wnd.getClassName() not in (
                    "Edit",
                    "LauncherTipWnd",
                    "Windows.UI.Input.InputSite.WindowClass",
                )
            if wnd.getProcessName() == "TE64.exe":
                return wnd.getClassName() == "SysListView32"
            return False

        @staticmethod
        def is_tablacus_viewmode(wnd: pyauto.Window) -> bool:
            return wnd.getProcessName() == "TE64.exe" and wnd.getClassName() == "SysListView32"

        @staticmethod
        def is_sumatra(wnd: pyauto.Window) -> bool:
            return wnd.getProcessName() == "SumatraPDF.exe"

        @classmethod
        def is_sumatra_viewmode(cls, wnd: pyauto.Window) -> bool:
            return cls.is_sumatra(wnd) and wnd.getClassName() != "Edit"

        @classmethod
        def is_sumatra_inputmode(cls, wnd: pyauto.Window) -> bool:
            return cls.is_sumatra(wnd) and wnd.getClassName() == "Edit"

    # keymap working on any window
    keymap_global = keymap.defineWindowKeymap(check_func=CheckWnd.is_global_target)

    KEYHAC_SANDS = False

    if KEYHAC_SANDS:
        keymap.replaceKey("Space", "RShift")
        keymap.replaceKey("RShift", "LShift")
        keymap_global["O-RShift"] = "Space"

    # keyboard macro
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["S-U0-0"] = keymap.command_RecordClear
    keymap_global["U1-0"] = keymap.command_RecordPlay
    keymap_global["U1-F4"] = keymap.command_RecordPlay

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
            # delete to bol / eol
            "S-U0-B": ("S-Home", "Delete"),
            "S-U0-D": ("S-End", "Delete"),
            # escape
            "O-(235)": ("Esc"),
            "U0-X": ("Esc"),
            # close window
            "LC-Q": ("A-F4"),
            # select last character
            "U1-U": ("LS-Left"),
            # SKK: contbvert to first suggestion
            "U0-Tab": ("C-N", "C-J"),
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
            # as katakana suggestion
            "C-U0-I": ("C-I", "S-Slash", "C-M"),
            # Context menu
            "U0-C": ("Apps"),
            "S-U0-C": ("S-F10"),
            # indent / outdent
            "U1-Tab": ("LC-CloseBracket"),
            "LS-U1-Tab": ("LC-OpenBracket"),
            # rename
            "U0-N": ("F2", "Right"),
            "S-U0-N": ("F2", "C-Home"),
            "C-U0-N": ("F2"),
            # print
            "F1": ("C-P"),
            "U1-F1": ("F1"),
            "Insert": (lambda: None),
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

    VIRTUAL_FINGER = VirtualFinger(keymap, 10)
    VIRTUAL_FINGER_QUICK = VirtualFinger(keymap, 0)

    class ImeControl:
        kana_key = "C-J"
        latin_key = "S-L"
        cancel_key = "C-G"
        reconv_key = "LWin-Slash"
        abbrev_key = "Slash"

        def __init__(
            self,
            keymap: Keymap,
        ) -> None:
            self._keymap = keymap
            self._finger = VirtualFinger(self._keymap, 0)

        def get_status(self) -> int:
            return self._keymap.getWindow().getImeStatus()

        def set_status(self, mode: int) -> None:
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

        def reconvert_with_skk(self) -> None:
            self.enable_skk()
            self._finger.type_keys(self.reconv_key, self.cancel_key)

        def disable(self) -> None:
            self.set_status(0)

    IME_CONTROL = ImeControl(keymap)

    class ClipHandler:
        @staticmethod
        def get_string() -> str:
            return getClipboardText() or ""

        @staticmethod
        def set_string(s: str) -> None:
            setClipboardText(str(s))

        @classmethod
        def paste(cls, s: str, format_func: Union[Callable, None] = None) -> None:
            if format_func is not None:
                cls.set_string(format_func(s))
            else:
                cls.set_string(s)
            VIRTUAL_FINGER_QUICK.type_keys("C-V")

        @classmethod
        def paste_current(cls, format_func: Union[Callable, None] = None) -> None:
            cls.paste(cls.get_string(), format_func)

        @classmethod
        def copy_string(cls) -> str:
            cls.set_string("")
            VIRTUAL_FINGER_QUICK.type_keys("C-Insert")
            interval = 10
            timeout = interval * 20
            while timeout > 0:
                if s := cls.get_string():
                    return s
                delay(interval)
                timeout -= interval
            return ""

        @classmethod
        def append(cls) -> None:
            org = cls.get_string()
            cb = cls.copy_string()
            cls.set_string(org + cb)

    class LazyFunc:
        def __init__(self, func: Callable) -> None:
            self._keymap = keymap

            def _wrapper() -> None:
                self._keymap.hookCall(func)

            self._func = _wrapper

        def defer(self, msec: int = 20) -> Callable:
            def _executer() -> None:
                # use delayedCall in ckit => https://github.com/crftwr/ckit/blob/2ea84f986f9b46a3e7f7b20a457d979ec9be2d33/ckitcore/ckitcore.cpp#L1998
                self._keymap.delayedCall(self._func, msec)

            return _executer

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

        def invoke(self, *sequence) -> Callable:
            def _input() -> None:
                self._control.disable()
                self._finger.type_smart(*sequence)
                if self._recover_ime:
                    self._control.enable()

            return LazyFunc(_input).defer(self._defer_msec)

    MILD_PUNCHER = KeyPuncher(keymap, defer_msec=20)
    SOFT_PUNCHER = KeyPuncher(keymap, defer_msec=50)


    ################################
    # release CapsLock on reload
    ################################

    if pyauto.Input.getKeyState(VK_CAPITAL):
        VIRTUAL_FINGER_QUICK.type_keys("LS-CapsLock")
        print("released CapsLock.")

    ################################
    # custom hotkey
    ################################

    # append clipboard
    keymap_global["LC-U0-C"] = LazyFunc(ClipHandler.append).defer()

    # listup Window
    keymap_global["U0-W"] = LazyFunc(lambda: VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")).defer()

    # ime: Japanese / Foreign
    keymap_global["U1-J"] = IME_CONTROL.enable_skk
    keymap_global["U0-F"] = IME_CONTROL.to_skk_latin
    keymap_global["S-U0-F"] = IME_CONTROL.enable_skk
    keymap_global["S-U1-J"] = IME_CONTROL.to_skk_latin
    keymap_global["U1-I"] = IME_CONTROL.reconvert_with_skk
    keymap_global["U1-R"] = IME_CONTROL.reconvert_with_skk
    keymap_global["U0-R"] = IME_CONTROL.reconvert_with_skk
    keymap_global["O-(236)"] = IME_CONTROL.to_skk_abbrev

    # paste as plaintext
    keymap_global["U0-V"] = LazyFunc(ClipHandler().paste_current).defer()

    # paste as plaintext (with trimming removable whitespaces)
    class StrCleaner:
        @staticmethod
        def clear_space(s: str) -> str:
            return s.strip().translate(str.maketrans("", "", "\u200b\u3000\u0009\u0020"))

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

            return LazyFunc(_paster).defer()

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

    # count chars
    def count_chars() -> None:
        cb = ClipHandler().copy_string()
        if cb:
            total = len(cb)
            lines = len(cb.strip().splitlines())
            net = len(StrCleaner.clear_space(cb.strip()))
            t = "total: {}(lines: {}), net: {}".format(total, lines, net)
            keymap.popBalloon("", t, 5000)

    keymap_global["LC-U1-C"] = LazyFunc(count_chars).defer()

    # wrap with quote mark
    def quote_selection() -> None:
        cb = ClipHandler().copy_string()
        if cb:
            ClipHandler().paste(cb, lambda x: ' "{}" '.format(x.strip()))

    keymap_global["LC-U0-Q"] = LazyFunc(quote_selection).defer()

    # paste with quote mark
    def paste_with_anchor(join_lines: bool = False) -> Callable:
        def _formatter(s: str) -> str:
            lines = s.strip().splitlines()
            if join_lines:
                return "> " + "".join([line.strip() for line in lines])
            return os.linesep.join(["> " + line for line in lines])

        def _paster() -> None:
            ClipHandler().paste_current(_formatter)

        return LazyFunc(_paster).defer()

    keymap_global["U1-Q"] = paste_with_anchor(False)
    keymap_global["C-U1-Q"] = paste_with_anchor(True)

    # open url in browser
    keymap_global["C-U0-O"] = LazyFunc(
        lambda: PathHandler(ClipHandler().copy_string().strip()).run()
    ).defer()

    # re-input selected string with skk
    def re_input_with_skk() -> None:
        selection = ClipHandler().copy_string()
        if selection:
            sequence = ["Minus" if c == "-" else c for c in StrCleaner.clear_space(selection)]
            IME_CONTROL.enable_skk()
            c = ord(sequence[0])
            if ord("A") <= c <= ord("Z") or ord("a") <= c <= ord("z"):
                sequence = ["LS-" + sequence[0]] + sequence[1:]
                VIRTUAL_FINGER_QUICK.type_smart(*sequence)
            else:
                IME_CONTROL.reconvert_with_skk()

    keymap_global["U1-Back"] = LazyFunc(re_input_with_skk).defer()
    keymap_global["U1-F9"] = LazyFunc(re_input_with_skk).defer()

    def moko(search_all: bool = False) -> Callable:
        exe_path = PathHandler(r"Personal\tools\bin\moko.exe", True)
        src_path = PathHandler(r"Personal\launch.yaml", True)

        def _launcher() -> None:
            exe_path.run(
                "-src={}".format(src_path.path),
                "-filer={}".format(KEYHAC_FILER),
                "-all={}".format(search_all),
                "-exclude=_obsolete,node_modules",
            )

        return LazyFunc(_launcher).defer()

    keymap_global["U1-Z"] = moko(False)
    keymap_global["LC-U1-Z"] = moko(True)

    ################################
    # config menu
    ################################

    class ConfigMenu:
        def __init__(self) -> None:
            self._keymap = keymap

        @staticmethod
        def read_config() -> str:
            return Path(getAppExePath(), "config.py").read_text("utf-8")

        def reload_config(self) -> None:
            self._keymap.configure()
            self._keymap.updateKeymap()
            ts = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
            print("\n{} reloaded config.py\n".format(ts))

        @staticmethod
        def open_repo() -> None:
            repo_path = PathHandler(r"Sync\develop\repo\keyhac", True)
            if repo_path.is_accessible():
                if KEYHAC_EDITOR == "notepad.exe":
                    print("notepad.exe cannot open directory. instead, open directory on explorer.")
                    repo_path.run()
                else:
                    PathHandler(KEYHAC_EDITOR).run(repo_path.path)
            else:
                print("cannot find path: '{}'".format(repo_path.path))

        @staticmethod
        def open_skk_config() -> None:
            skk_path = PathHandler(r"C:\Windows\System32\IME\IMCRVSKK\imcrvcnf.exe")
            skk_path.run()

        @staticmethod
        def open_skk_dir() -> None:
            skk_dir_path = PathHandler.resolve_user_profile(r"AppData\Roaming\CorvusSKK")
            PathHandler(KEYHAC_FILER).run(skk_dir_path)

        def apply(self, km: WindowKeymap) -> None:
            for key, func in {
                "R": self.reload_config,
                "E": self.open_repo,
                "P": lambda: ClipHandler().paste(self.read_config()),
                "S": self.open_skk_config,
                "S-S": self.open_skk_dir,
                "X": lambda: None,
            }.items():
                km[key] = LazyFunc(func).defer()

    keymap_global["LC-U0-X"] = keymap.defineMultiStrokeKeymap()
    ConfigMenu().apply(keymap_global["LC-U0-X"])

    keymap_global["U1-F12"] = ConfigMenu().reload_config

    ################################
    # class for position on monitor
    ################################

    class WndRect:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap
            self._rect = (0, 0, 0, 0)
            self.left, self.top, self.right, self.bottom = self._rect

        def set_rect(self, left: int, top: int, right: int, bottom: int) -> None:
            self._rect = (left, top, right, bottom)
            self.left, self.top, self.right, self.bottom = self._rect

        def check_rect(self, wnd: pyauto.Window) -> bool:
            return wnd.getRect() == self._rect

        def snap(self) -> None:
            wnd = self._keymap.getTopLevelWindow()
            if self.check_rect(wnd):
                wnd.maximize()
                return
            if wnd.isMaximized():
                wnd.restore()
                delay()
            trial_limit = 2
            while trial_limit > 0:
                wnd.setRect(list(self._rect))
                delay()
                if self.check_rect(wnd):
                    return
                trial_limit -= 1

        def get_vertical_half(self, upper: bool) -> list:
            half_height = int((self.bottom + self.top) / 2)
            r = list(self._rect)
            if upper:
                r[3] = half_height
            else:
                r[1] = half_height
            if 200 < abs(r[1] - r[3]):
                return r
            return []

        def get_horizontal_half(self, leftward: bool) -> list:
            half_width = int((self.right + self.left) / 2)
            r = list(self._rect)
            if leftward:
                r[2] = half_width
            else:
                r[0] = half_width
            if 300 < abs(r[0] - r[2]):
                return r
            return []

    class MonitorRect:
        def __init__(self, keymap: Keymap, rect: list) -> None:
            self._keymap = keymap
            self.left, self.top, self.right, self.bottom = rect
            self.max_width = self.right - self.left
            self.max_height = self.bottom - self.top
            self.possible_width = self.get_variation(self.max_width)
            self.possible_height = self.get_variation(self.max_height)
            self.area_mapping = {}

        def set_center_rect(self) -> None:
            d = {}
            for size, px in self.possible_width.items():
                lx = self.left + int((self.max_width - px) / 2)
                wr = WndRect(self._keymap)
                wr.set_rect(lx, self.top, lx + px, self.bottom)
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
                    wr = WndRect(self._keymap)
                    wr.set_rect(lx, self.top, lx + px, self.bottom)
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
                    wr = WndRect(self._keymap)
                    wr.set_rect(self.left, ty, self.right, ty + px)
                    d[size] = wr
                self.area_mapping[pos] = d

        @staticmethod
        def get_variation(max_size: int) -> dict:
            return {
                "small": int(max_size / 3),
                "middle": int(max_size / 2),
                "large": int(max_size * 2 / 3),
            }

    class CurrentMonitors:
        def __init__(self, keymap: Keymap) -> None:
            ms = []
            for mi in pyauto.Window.getMonitorInfo():
                mr = MonitorRect(keymap, mi[1])
                mr.set_center_rect()
                mr.set_horizontal_rect()
                mr.set_vertical_rect()
                if mi[2] == 1:  # main monitor
                    ms.insert(0, mr)
                else:
                    ms.append(mr)
            self._monitors = ms

        def monitors(self) -> list:
            return self._monitors

    ################################
    # set cursor position
    ################################

    keymap_global["U0-Up"] = keymap.MouseWheelCommand(+0.5)
    keymap_global["U0-Down"] = keymap.MouseWheelCommand(-0.5)
    keymap_global["U0-Left"] = keymap.MouseHorizontalWheelCommand(-0.5)
    keymap_global["U0-Right"] = keymap.MouseHorizontalWheelCommand(+0.5)

    class CursorPos:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap
            self.pos = []
            monitors = CurrentMonitors(keymap).monitors()
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

    keymap_global["U1-Enter"] = CURSOR_POS.snap_to_center
    keymap_global["O-RShift"] = CURSOR_POS.snap_to_center

    ################################
    # set window position
    ################################

    keymap_global["U1-L"] = "LWin-Right"
    keymap_global["U1-H"] = "LWin-Left"

    keymap_global["U1-M"] = keymap.defineMultiStrokeKeymap()
    keymap_global["U1-M"]["X"] = LazyFunc(lambda: keymap.getTopLevelWindow().maximize()).defer()

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
            monitors = CurrentMonitors(self._keymap).monitors()
            for mod_mntr, mntr_idx in self.monitor_dict.items():
                for mod_area, size in self.size_dict.items():
                    for key, pos in self.snap_key_dict.items():
                        if mntr_idx < len(monitors):
                            wnd_rect = monitors[mntr_idx].area_mapping[pos][size]
                            km[mod_mntr + mod_area + key] = LazyFunc(wnd_rect.snap).defer()

        def alloc_maximize(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, towards in mapping_dict.items():

                def _snap() -> None:
                    VIRTUAL_FINGER.type_keys("LShift-LWin-" + towards)
                    delay()
                    self._keymap.getTopLevelWindow().maximize()

                km[key] = LazyFunc(_snap).defer()

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
            def _snap() -> None:
                wnd = self._keymap.getTopLevelWindow()
                wr = WndRect(self._keymap)
                wr.set_rect(*wnd.getRect())
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

            return LazyFunc(_snap).defer()

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
            def _send() -> None:
                self._control.enable_skk()
                self._finger.type_smart(*sequence)

            return _send

        def under_latinmode(self, *sequence) -> Callable:
            def _send() -> None:
                self._control.to_skk_latin()
                self._finger.type_smart(*sequence)

            return _send

    SIMPLE_SKK = SimpleSKK(keymap)

    # select-to-left with ime control
    keymap_global["U1-B"] = SIMPLE_SKK.under_kanamode("S-Left")
    keymap_global["LS-U1-B"] = SIMPLE_SKK.under_latinmode("S-Left")
    keymap_global["U1-Space"] = SIMPLE_SKK.under_kanamode("C-S-Left")
    keymap_global["LS-U1-Space"] = SIMPLE_SKK.under_latinmode("C-S-Left")
    keymap_global["U1-N"] = SIMPLE_SKK.under_kanamode("S-Left", ImeControl.abbrev_key)

    class SKK:
        def __init__(self, keymap: Keymap, finish_with_kanamode: bool = True) -> None:
            self._base_skk = SimpleSKK(keymap)
            self._finish_with_kanamode = finish_with_kanamode

        def send(self, *sequence) -> Callable:
            if self._finish_with_kanamode:
                sequence = list(sequence) + [ImeControl.kana_key]
            return self._base_skk.under_latinmode(*sequence)

        def send_pair(self, pair: list) -> Callable:
            _, suffix = pair
            sequence = pair + ["Left"] * len(suffix)
            if self._finish_with_kanamode:
                sequence = sequence + [ImeControl.kana_key]
            return self._base_skk.under_latinmode(*sequence)

        def apply(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, sent in mapping_dict.items():
                km[key] = self.send(sent)

        def apply_pair(self, km: WindowKeymap, mapping_dict: dict) -> None:
            for key, sent in mapping_dict.items():
                km[key] = self.send_pair(sent)

    SKK_TO_KANAMODE = SKK(keymap, True)
    SKK_TO_LATINMODE = SKK(keymap, False)

    # markdown list
    keymap_global["S-U0-8"] = SKK_TO_KANAMODE.send("- ")
    keymap_global["U1-1"] = SKK_TO_KANAMODE.send("1. ")

    SKK_TO_KANAMODE.apply(
        keymap_global,
        {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Minus": "\u3000\u2015\u2015",
            "S-U0-Period": "\uff0e",  # FULLWIDTH FULL STOP
            "S-U0-U": "S-BackSlash",
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
            "U0-T": ["\u3014", "\u3015"],  # TORTOISE SHELL BRACKET 〔〕
            "U1-8": ["\uff08", "\uff09"],  # FULLWIDTH PARENTHESIS （）
            "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
        },
    )

    SKK_TO_LATINMODE.apply(
        keymap_global,
        {
            "U0-1": "S-1",
            "U0-4": "$_",
            "U0-Colon": "Colon",
            "U0-Comma": "Comma",
            "U0-Period": "Period",
            "U0-Slash": "Slash",
            "U1-Minus": "Minus",
            "U0-SemiColon": "SemiColon",
            "U1-SemiColon": "+:",
        },
    )
    SKK_TO_LATINMODE.apply_pair(
        keymap_global,
        {
            "U0-CloseBracket": ["[", "]"],
            "U1-9": ["(", ")"],
            "U1-CloseBracket": ["{", "}"],
            "U0-Caret": ["~~", "~~"],
        },
    )

    class DateInput:
        def __init__(self, km: WindowKeymap) -> None:
            self._keymap = km

        @staticmethod
        def invoke(fmt: str, after_mode_is_kana: bool = False) -> Callable:
            def _input() -> None:
                d = datetime.datetime.today()
                seq = [c for c in d.strftime(fmt)]
                if after_mode_is_kana:
                    SKK_TO_KANAMODE.send(*seq)()
                else:
                    SKK_TO_LATINMODE.send(*seq)()

            return LazyFunc(_input).defer()

        def apply(self) -> None:
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
                self._keymap[key] = self.invoke(*params)

    keymap_global["U1-D"] = keymap.defineMultiStrokeKeymap()
    DateInput(keymap_global["U1-D"]).apply()

    ################################
    # pseudo espanso
    ################################

    def combo_mapper(
        root_map: Union[Callable, dict], keys: list, func: Callable
    ) -> Union[Callable, dict, None]:
        if callable(root_map):
            return root_map
        if len(keys) == 1:
            root_map[keys[0]] = func
            return root_map
        if len(keys) < 1:
            return None
        head = keys[0]
        rest = keys[1:]
        try:
            root_map[head] = combo_mapper(root_map[head], rest, func)
        except:
            sub_map = keymap.defineMultiStrokeKeymap()
            root_map[head] = combo_mapper(sub_map, rest, func)
        return root_map

    class KeyCombo:
        def __init__(self, keymap: Keymap) -> None:
            self._keymap = keymap
            self.mapping = self._keymap.defineMultiStrokeKeymap()
            for combo, stroke in {
                "X,X": [".txt"],
                "X,M": [".md"],
                "Minus,H": ["\u2010"],
                "Minus,M": ["\u2014"],
                "Minus,N": ["\u2013"],
                "Minus,S": ["\u2212"],
                "Minus,D": ["\u30a0"],
                "R,B": [r"[\[［].+?[\]］]"],
                "R,P": [r"[\(（].+?[\)）]"],
                "C-R,S-Right": ["(?!)", "Left"],
                "C-R,Right": ["(?=)", "Left"],
                "C-R,S-Left": ["(?<!)", "Left"],
                "C-R,Left": ["(?<=)", "Left"],
                "A,E": ["\u00e9"],
                "A,S-E": ["\u00c9"],
                "C-A,E": ["\u00e8"],
                "C-A,S-E": ["\u00c8"],
                "U,S-A": ["\u00c4"],
                "U,A": ["\u00e4"],
                "U,S-O": ["\u00d6"],
                "U,O": ["\u00f6"],
                "U,S-U": ["\u00dc"],
                "U,U": ["\u00fc"],
            }.items():
                keys = combo.split(",")
                self.mapping = combo_mapper(self.mapping, keys, SKK_TO_LATINMODE.send(*stroke))

            for combo, stroke in {
                "Minus,F": ["\uff0d"],
                "F,G": ["\u3013\u3013"],
                "F,0": ["\u25cf\u25cf"],
                "F,4": ["\u25a0\u25a0"],
                "M,1": ["# "],
                "M,2": ["## "],
                "M,3": ["### "],
                "M,4": ["#### "],
                "M,5": ["##### "],
                "M,6": ["###### "],
            }.items():
                keys = combo.split(",")
                self.mapping = combo_mapper(self.mapping, keys, SKK_TO_KANAMODE.send(*stroke))

    keymap_global["U1-X"] = KeyCombo(keymap).mapping

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

        def get_mapping(self) -> dict:
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

            self._mapping = _mapper.get_mapping()

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
            def _search() -> None:
                s = ClipHandler().copy_string()
                query = SearchQuery(s)
                query.fix_kangxi()
                query.remove_honorific()
                query.remove_editorial_style()
                if strip_hiragana:
                    query.remove_hiragana()
                PathHandler(uri.format(query.encode(strict))).run()

            return LazyFunc(_search).defer()

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
            "B": "https://www.google.com/search?q=site%3Abooks.or.jp%20{}",
            "C": "https://ci.nii.ac.jp/books/search?q={}",
            "D": "https://duckduckgo.com/?q={}",
            "E": "http://webcatplus.nii.ac.jp/pro/?q={}",
            "G": "http://www.google.co.jp/search?q={}",
            "H": "https://www.hanmoto.com/bd/search/order/desc/title/{}",
            "I": "https://www.google.com/search?tbm=isch&q={}",
            "J": "https://eow.alc.co.jp/search?q={}",
            "K": "https://www.kinokuniya.co.jp/disp/CSfDispListPage_001.jsp?qs=true&ptk=01&q={}",
            "M": "https://www.google.co.jp/maps/search/{}",
            "N": "https://iss.ndl.go.jp/books?any={}",
            "P": "https://wordpress.org/openverse/search/?q={}",
            "R": "https://researchmap.jp/researchers?q={}",
            "S": "https://scholar.google.co.jp/scholar?as_vis=1&q={}",
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
            interval = 10
            if target.isMinimized():
                target.restore()
                delay(interval)
            timeout = interval * 10
            while timeout > 0:
                try:
                    target.setForeground()
                    if pyauto.Window.getForeground() == target:
                        target.setForeground(True)
                        return True
                except:
                    return False
                delay(interval)
                timeout -= interval
            return False

        def invoke(self, exe_name: str, class_name: str = "", exe_path: str = "") -> Callable:
            scanner = WndScanner(exe_name, class_name)

            def _executer() -> None:
                scanner.scan()
                if scanner.found:
                    if scanner.found == self._keymap.getWindow() or not self.activate_wnd(
                        scanner.found
                    ):
                        VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")
                else:
                    if exe_path:
                        PathHandler(exe_path).run()

            return LazyFunc(_executer).defer(80)

        def apply(self, wnd_keymap: WindowKeymap, remap_table: dict = {}) -> None:
            for key, params in remap_table.items():
                wnd_keymap[key] = self.invoke(*params)

    PSEUDO_CUTEEXEC = PseudoCuteExec(keymap)

    PSEUDO_CUTEEXEC.apply(
        keymap_global,
        {
            "U1-T": ("TE64.exe", "TablacusExplorer"),
            "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
            "LC-U1-M": (
                "Mery.exe",
                "TChildForm",
                PathHandler.resolve_user_profile(r"AppData\Local\Programs\Mery\Mery.exe"),
            ),
            "LC-U1-N": (
                "notepad.exe",
                "Notepad",
                r"C:\Windows\System32\notepad.exe",
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
                PathHandler.resolve_user_profile(r"AppData\Local\Vivaldi\Application\vivaldi.exe"),
            ),
            "S": (
                "slack.exe",
                "Chrome_WidgetWin_1",
                PathHandler.resolve_user_profile(r"AppData\Local\slack\slack.exe"),
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
                PathHandler.resolve_user_profile(r"scoop\apps\ksnip\current\ksnip.exe"),
            ),
            "O": ("Obsidian.exe", "Chrome_WidgetWin_1"),
            "P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
            "C-P": ("powerpnt.exe", "PPTFrameClass"),
            "E": ("EXCEL.EXE", "XLMAIN"),
            "W": ("WINWORD.EXE", "OpusApp"),
            "V": ("Code.exe", "Chrome_WidgetWin_1", KEYHAC_EDITOR),
            "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1"),
            "T": ("TE64.exe", "", KEYHAC_FILER),
            "M": (
                "Mery.exe",
                "TChildForm",
                PathHandler.resolve_user_profile(r"AppData\Local\Programs\Mery\Mery.exe"),
            ),
            "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
        },
    )

    keymap_global["LS-LC-U1-M"] = PathHandler(r"Personal\draft.txt", True).run

    # invoke specific filer
    def invoke_filer(dir_path: str) -> Callable:
        def _invoker() -> None:
            if PathHandler(dir_path).is_accessible():
                PathHandler(KEYHAC_FILER).run(dir_path)
            else:
                print("invalid-path: '{}'".format(dir_path))

        return LazyFunc(_invoker).defer()

    keymap_global["U1-F"] = keymap.defineMultiStrokeKeymap()
    keymap_global["U1-F"]["D"] = invoke_filer(PathHandler.resolve_user_profile(r"Desktop"))
    keymap_global["U1-F"]["S"] = invoke_filer(r"X:\scan")

    def invoke_terminal() -> None:
        scanner = WndScanner("wezterm-gui.exe", "org.wezfurlong.wezterm")
        scanner.scan()
        if scanner.found:
            if PSEUDO_CUTEEXEC.activate_wnd(scanner.found):
                return
        PathHandler(r"scoop\apps\wezterm\current\wezterm-gui.exe", True).run()

    keymap_global["LC-AtMark"] = LazyFunc(invoke_terminal).defer()

    def search_on_browser() -> None:
        if keymap.getWindow().getProcessName() == DEFAULT_BROWSER.get_exe_name():
            VIRTUAL_FINGER.type_keys("C-T")
        else:
            scanner = WndScanner(DEFAULT_BROWSER.get_exe_name(), DEFAULT_BROWSER.get_wnd_class())
            scanner.scan()
            if scanner.found:
                if PSEUDO_CUTEEXEC.activate_wnd(scanner.found):
                    VIRTUAL_FINGER.type_keys("C-T")
                else:
                    VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")
            else:
                PathHandler("https://duckduckgo.com").run()

    keymap_global["U0-Q"] = LazyFunc(search_on_browser).defer(100)

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=CheckWnd.is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"

    def focus_main_pane() -> None:
        wnd = keymap.getWindow()
        if wnd.getProcessName() == "chrome.exe":
            VIRTUAL_FINGER.type_keys("S-A-B", "F6")
        elif wnd.getProcessName() == "firefox.exe":
            VIRTUAL_FINGER.type_keys("C-L", "F6")

    keymap_browser["U0-F6"] = focus_main_pane

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None

    # slack
    keymap_slack = keymap.defineWindowKeymap(exe_name="slack.exe", class_name="Chrome_WidgetWin_1")
    keymap_slack["F3"] = SIMPLE_SKK.under_latinmode("C-K")
    keymap_slack["C-E"] = SIMPLE_SKK.under_latinmode("C-K")
    keymap_slack["F1"] = SIMPLE_SKK.under_latinmode("+:")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")

    def remap_vscode(keys: list, km: WindowKeymap) -> Callable:
        for key in keys:
            km[key] = MILD_PUNCHER.invoke(key)

    remap_vscode(
        [
            "C-F",
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
    keymap_sumatra = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra)

    keymap_sumatra["C-S-F"] = SIMPLE_SKK.under_latinmode("C-Home", "Esc", "C-F", "C-J")
    keymap_sumatra["LA-X"] = SIMPLE_SKK.under_latinmode("C-Home", "Esc", "C-F", "C-J")

    keymap_sumatra_inputmode = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra_inputmode)

    def sumatra_change_tab(km: WindowKeymap) -> None:
        for key in ["C-Tab", "C-S-Tab"]:
            km[key] = "Esc", key

    sumatra_change_tab(keymap_sumatra_inputmode)

    keymap_sumatra_viewmode = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra_viewmode)

    def sumatra_view_key(km: WindowKeymap) -> None:
        for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            km[key] = SIMPLE_SKK.under_latinmode(key)

    sumatra_view_key(keymap_sumatra_viewmode)

    keymap_sumatra_viewmode["F"] = SIMPLE_SKK.under_kanamode("C-F")
    keymap_sumatra_viewmode["H"] = "C-S-Tab"
    keymap_sumatra_viewmode["L"] = "C-Tab"
    keymap_sumatra_viewmode["X"] = SIMPLE_SKK.under_latinmode("C-Home", "Esc", "C-F", "C-J")

    def office_to_pdf(km: WindowKeymap, key: str = "F11") -> None:
        km[key] = "A-F", "E", "P", "A"

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")
    office_to_pdf(keymap_word)

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    office_to_pdf(keymap_ppt)

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

    # filer
    keymap_filer = keymap.defineWindowKeymap(check_func=CheckWnd.is_filer_viewmode)
    KeyAllocator(
        {
            "C-SemiColon": ("C-Add"),
            "C-L": ("A-D", "C-C"),
        }
    ).apply(keymap_filer)

    ################################
    # popup clipboard menu
    ################################

    class CharWidth:
        full_letters = "\uff41\uff42\uff43\uff44\uff45\uff46\uff47\uff48\uff49\uff4a\uff4b\uff4c\uff4d\uff4e\uff4f\uff50\uff51\uff52\uff53\uff54\uff55\uff56\uff57\uff58\uff59\uff5a\uff21\uff22\uff23\uff24\uff25\uff26\uff27\uff28\uff29\uff2a\uff2b\uff2c\uff2d\uff2e\uff2f\uff30\uff31\uff32\uff33\uff34\uff35\uff36\uff37\uff38\uff39\uff3a\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19\uff0d"
        half_letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"

        @classmethod
        def to_half_letter(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.full_letters, cls.half_letters))

        @classmethod
        def to_full_letter(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.half_letters, cls.full_letters))

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
            return (d.strftime("%Y年%m月%d日（{}） {}%H:%M～")).format(week, ampm)

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
                print("ERROR: lack of lines.")
                return ""
            due = cls.get_time(lines[3])
            if len(due) < 1:
                print("ERROR: could not parse due date.")
                return ""
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

    class ClipboardMenu:
        def __init__(self) -> None:
            pass

        @staticmethod
        def format_cb(func: Callable) -> Callable:
            def _formatter() -> str:
                cb = ClipHandler.get_string()
                if cb:
                    return func(cb)

            return _formatter

        @staticmethod
        def replace_cb(search: str, replace_to: str) -> Callable:
            reg = re.compile(search)

            def _replacer() -> str:
                cb = ClipHandler.get_string()
                if cb:
                    return reg.sub(replace_to, cb)

            return _replacer

        @staticmethod
        def catanate_file_content(s: str) -> str:
            if PathHandler(s).is_accessible():
                return Path(s).read_text("utf-8")
            return ""

        @staticmethod
        def as_codeblock(s: str) -> str:
            return "\n".join(["```", s, "```"])

        @staticmethod
        def skip_blank_line(s: str) -> str:
            lines = s.strip().splitlines()
            return os.linesep.join([l for l in lines if l.strip()])

        @staticmethod
        def to_double_bracket(s: str) -> str:
            reg = re.compile(r"[\u300c\u300d]")

            def _replacer(mo: re.Match) -> str:
                if mo.group(0) == "\u300c":
                    return "\u300e"
                return "\u300f"

            return reg.sub(_replacer, s)

        @staticmethod
        def split_postalcode(s: str) -> str:
            reg = re.compile(r"(\d{3}).(\d{4})[\s\r\n]*(.+$)")
            hankaku = CharWidth().to_half_letter(s.strip().strip("\u3012"))
            m = reg.match(hankaku)
            if m:
                return "{}-{}\t{}".format(m.group(1), m.group(2), m.group(3))
            return s

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

        @classmethod
        def get_menu_noise_reduction(cls) -> list:
            return [
                (" Remove: - Blank lines ", cls.format_cb(cls.skip_blank_line)),
                (
                    "         - Inside Paren ",
                    cls.replace_cb(r"[\uff08\u0028].+?[\uff09\u0029]", ""),
                ),
                ("         - Line-break ", cls.replace_cb(r"\r?\n", "")),
                ("         - Non-digit-char ", cls.replace_cb(r"[^\d]", "")),
                ("         - Quotations ", cls.replace_cb(r"[\u0022\u0027]", "")),
                (" Fix: - Dumb Quotation ", cls.format_cb(cls.fix_dumb_quotation)),
                ("      - MSWord-Bullet ", cls.replace_cb(r"\uf09f\u0009", "\u30fb")),
                ("      - KANGXI RADICALS ", cls.format_cb(KangxiRadicals().fix)),
                ("      - To-Double-Bracket ", cls.format_cb(cls.to_double_bracket)),
            ]

        @classmethod
        def get_menu_transform(cls) -> list:
            return [
                (" Transform: => A-Z/0-9 ", cls.format_cb(CharWidth().to_half_letter)),
                (
                    "            => \uff21-\uff3a/\uff10-\uff19 ",
                    cls.format_cb(CharWidth().to_full_letter),
                ),
                ("            => abc ", lambda: ClipHandler.get_string().lower()),
                ("            => ABC ", lambda: ClipHandler.get_string().upper()),
                (" Comma: - Curly (\uff0c) ", cls.replace_cb(r"\u3001", "\uff0c")),
                ("        - Straight (\u3001) ", cls.replace_cb(r"\uff0c", "\u3001")),
            ]

        @classmethod
        def get_menu_other(cls) -> list:
            return [
                (" As md-codeblock ", cls.format_cb(cls.as_codeblock)),
                (" Cat local file ", cls.format_cb(cls.catanate_file_content)),
                (" Postalcode | Address ", cls.format_cb(cls.split_postalcode)),
                (" URL: - Decode ", cls.format_cb(cls.decode_url)),
                ("      - Encode ", cls.format_cb(cls.encode_url)),
                (
                    "      - Shorten Amazon ",
                    cls.replace_cb(
                        r"^.+amazon\.co\.jp/.+dp/(.{10}).*", r"https://www.amazon.jp/dp/\1"
                    ),
                ),
                (" Zoom invitation ", cls.format_cb(Zoom().format)),
            ]

        @classmethod
        def apply(cls, km: WindowKeymap) -> None:
            for title, menu in {
                "Noise-Reduction": cls.get_menu_noise_reduction(),
                "Transform Alphabet / Punctuation": cls.get_menu_transform(),
                "Others": cls.get_menu_other(),
            }.items():
                m = menu + [("---------------- EXIT ----------------", lambda: None)]
                km.cblisters += [(title, cblister_FixedPhrase(m))]

    ClipboardMenu().apply(keymap)
    keymap_global["LC-LS-X"] = LazyFunc(keymap.command_ClipboardList).defer(msec=100)
