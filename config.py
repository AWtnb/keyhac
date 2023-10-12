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


def configure(keymap):
    ################################
    # general setting
    ################################

    class PathInfo:
        def __init__(self, s: str) -> None:
            self.path = s
            self.isAccessible = self.path.startswith("http") or Path(self.path).exists()

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
            if self.isAccessible:
                keymap.ShellExecuteCommand(None, self.path, self.args_to_param(args), None)()
            else:
                print("invalid-path: '{}'".format(self.path))

    class UserPath:
        user_prof = os.environ.get("USERPROFILE") or ""

        @classmethod
        def resolve(cls, rel: str = "") -> PathInfo:
            fullpath = str(Path(cls.user_prof, rel))
            return PathInfo(fullpath)

        @classmethod
        def mask_user_name(cls, path: str) -> str:
            masked = str(Path(cls.user_prof).parent) + r"\%USERNAME%"
            return masked + path[len(cls.user_prof) :]

        @classmethod
        def get_filer(cls) -> str:
            tablacus = cls.resolve(r"Sync\portable_app\tablacus\TE64.exe")
            if tablacus.isAccessible:
                return tablacus.path
            return "explorer.exe"

        @classmethod
        def get_editor(cls) -> str:
            vscode = cls.resolve(r"scoop\apps\vscode\current\Code.exe")
            if vscode.isAccessible:
                return vscode.path
            return "notepad.exe"

    keymap.editor = UserPath().get_editor()

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

    # keyboard macro
    keymap_global["U0-R"] = keymap.command_RecordPlay
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["U1-0"] = keymap.command_RecordClear

    # combination with modifier key
    class CoreKeys:
        mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
        key_status = ("D-", "U-")

        @classmethod
        def cursor_keys(cls, km: Keymap) -> None:
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
        def ignore_capslock(cls, km: Keymap) -> None:
            for stat in cls.key_status:
                for mod_key in cls.mod_keys:
                    km[mod_key + stat + "Capslock"] = lambda: None

        @classmethod
        def ignore_kanakey(cls, km: Keymap) -> None:
            for stat in cls.key_status:
                for mod_key in cls.mod_keys:
                    for vk in list(range(124, 136)) + list(range(240, 243)) + list(range(245, 254)):
                        km[mod_key + stat + str(vk)] = lambda: None

    CoreKeys().cursor_keys(keymap_global)
    CoreKeys().ignore_capslock(keymap_global)
    CoreKeys().ignore_kanakey(keymap_global)

    class KeyAllocator:
        def __init__(self, mapping_dict: dict) -> None:
            self.dict = mapping_dict

        def apply(self, km: Keymap):
            for key, value in self.dict.items():
                km[key] = value

    KeyAllocator(
        {
            # delete to bol / eol
            "S-U0-B": ("S-Home", "Delete"),
            "S-U0-D": ("S-End", "Delete"),
            # escape
            "O-(235)": ("Esc"),
            "U0-X": ("Esc"),
            # select first suggestion
            "U0-Tab": ("Down", "Enter"),
            # confirm and move cursor
            "U0-M": ("Enter", "Right"),
            # line selection
            "U1-A": ("End", "S-Home"),
            # punctuation
            "U0-Enter": ("Period"),
            "LS-U0-Enter": ("Comma"),
            "LC-U0-Enter": ("Slash"),
            "U1-B": ("Minus"),
            "U0-Z": ("Minus"),
            # Re-convert
            "S-(236)": ("LWin-Slash"),
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

    ################################
    # functions for custom hotkey
    ################################

    def delay(msec: int = 50) -> None:
        time.sleep(msec / 1000)

    class VirtualFinger:
        def __init__(self, inter_stroke_pause: int = 10) -> None:
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
            self._keymap.setInput_Modifier(0)
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

    VIRTUAL_FINGER = VirtualFinger(10)
    VIRTUAL_FINGER_QUICK = VirtualFinger(0)

    class ImeControl:
        def __init__(self) -> None:
            self._keymap = keymap

        def get_status(self) -> pyauto.Window:
            return self._keymap.getWindow().getImeStatus()

        def set_status(self, mode: int) -> None:
            if self.get_status() != mode:
                VIRTUAL_FINGER_QUICK.type_keys("(243)")
                delay(10)

        def is_enabled(self) -> bool:
            return self.get_status() == 1

        def enable(self) -> None:
            self.set_status(1)

        def disable(self) -> None:
            self.set_status(0)

    IME_CONTROL = ImeControl()

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
            VIRTUAL_FINGER_QUICK.type_keys("S-Insert")

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
            recover_ime: bool = False,
            inter_stroke_pause: int = 0,
            defer_msec: int = 0,
        ) -> None:
            self._recover_ime = recover_ime
            self._inter_stroke_pause = inter_stroke_pause
            self._defer_msec = defer_msec

        def invoke(self, *sequence) -> Callable:
            vf = VirtualFinger(self._inter_stroke_pause)

            def _input() -> None:
                IME_CONTROL.disable()
                vf.type_smart(*sequence)
                if self._recover_ime:
                    IME_CONTROL.enable()

            return LazyFunc(_input).defer(self._defer_msec)

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

    # switch window
    keymap_global["U1-Tab"] = KeyPuncher(defer_msec=40).invoke("Alt-Tab")

    # ime dict tool
    keymap_global["U0-F7"] = LazyFunc(lambda: PathInfo(r"C:\Program Files (x86)\Google\Google Japanese Input\GoogleIMEJaTool.exe").run("--mode=word_register_dialog")).defer()

    # listup Window
    keymap_global["U0-W"] = LazyFunc(lambda: VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")).defer()

    # ime: Japanese / Foreign
    keymap_global["U1-J"] = IME_CONTROL.enable
    keymap_global["U0-F"] = IME_CONTROL.disable
    keymap_global["S-U0-F"] = IME_CONTROL.enable
    keymap_global["S-U1-J"] = IME_CONTROL.disable

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
        def apply(cls, km: Keymap, custom_key: str) -> None:
            for mod_ctrl, remove_white in {
                "": False,
                "LC-": True,
            }.items():
                for mod_shift, include_linebreak in {
                    "": False,
                    "LS-": True,
                }.items():
                    km[mod_ctrl + mod_shift + custom_key] = cls.invoke(remove_white, include_linebreak)

    StrCleaner().apply(keymap_global, "U1-V")

    # select last word with ime
    def select_last_word() -> None:
        VIRTUAL_FINGER.type_keys("C-S-Left")
        IME_CONTROL.enable()

    keymap_global["U1-Space"] = select_last_word

    # Non-convert
    def as_alphabet(recover_ime: bool = False) -> Callable:
        keys = ["F10"]
        if recover_ime:
            keys.append("Enter")
        else:
            keys.append("(243)")

        def _sender():
            if keymap.getWindow().getImeStatus() == 1:
                VIRTUAL_FINGER.type_keys(*keys)

        return _sender

    keymap_global["U1-N"] = as_alphabet(False)
    keymap_global["S-U1-N"] = as_alphabet(True)

    def as_titled_alphabet() -> None:
        if keymap.getWindow().getImeStatus() == 1:
            VIRTUAL_FINGER.type_keys("F10", "F10", "F10", "C-Space")

    keymap_global["LA-U1-N"] = as_titled_alphabet

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
    keymap_global["C-U0-O"] = LazyFunc(lambda: PathInfo(ClipHandler().copy_string().strip()).run()).defer()

    # re-input selected string with ime
    def re_input_with_ime() -> None:
        selection = ClipHandler().copy_string()
        if selection:
            sequence = ["Minus" if c == "-" else c for c in StrCleaner.clear_space(selection)]
            IME_CONTROL.enable()
            VIRTUAL_FINGER_QUICK.type_smart(*sequence)

    keymap_global["U1-I"] = LazyFunc(re_input_with_ime).defer()

    def moko(search_all: bool = False) -> Callable:
        exe_path = r"C:\Personal\tools\bin\moko.exe"

        def _launcher() -> None:
            PathInfo(exe_path).run(
                r"-src=C:\Personal\launch.yaml",
                "-filer={}".format(UserPath().get_filer()),
                "-all={}".format(search_all),
                "-exclude=_obsolete,node_modules",
            )

        return LazyFunc(_launcher).defer()

    keymap_global["U1-Z"] = moko(False)
    keymap_global["LC-U1-Z"] = moko(True)

    def screenshot() -> None:
        ksnip_path = UserPath().resolve(r"scoop\apps\ksnip\current\ksnip.exe")
        if ksnip_path.isAccessible:
            ksnip_path.run()
        else:
            VIRTUAL_FINGER.type_keys("Lwin-S-S")

    keymap_global["C-U1-S"] = screenshot

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
            repo_path = UserPath().resolve(r"Sync\develop\repo\keyhac")
            if repo_path.isAccessible:
                editor_path = UserPath().get_editor()
                if editor_path == "notepad.exe":
                    print("notepad.exe cannot open directory. instead, open directory on explorer.")
                    PathInfo(repo_path.path).run()
                else:
                    PathInfo(editor_path).run(repo_path.path)
            else:
                print("cannot find path: '{}'".format(repo_path.path))

        def apply(self, km: Keymap) -> None:
            for key, func in {
                "R": self.reload_config,
                "E": self.open_repo,
                "P": lambda: ClipHandler().paste(self.read_config()),
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
        def __init__(self, left: int, top: int, right: int, bottom: int) -> None:
            self._rect = (left, top, right, bottom)
            self.left, self.top, self.right, self.bottom = self._rect

        def check_rect(self, wnd: pyauto.Window) -> bool:
            return wnd.getRect() == self._rect

        def snap(self) -> None:
            wnd = keymap.getTopLevelWindow()
            if self.check_rect(wnd):
                wnd.maximize()
                return
            if wnd.isMaximized():
                wnd.restore()
                delay()
            trial_limit = 2
            while trial_limit > 0:
                wnd.setRect(list(self._rect))
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
        def __init__(self, rect: list, is_primary: int) -> None:
            self.is_primary = is_primary  # 0 or 1
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
                d[size] = WndRect(lx, self.top, lx + px, self.bottom)
            self.area_mapping["center"] = d

        def set_horizontal_rect(self) -> None:
            for pos in ("left", "right"):
                d = {}
                for size, px in self.possible_width.items():
                    if pos == "right":
                        lx = self.right - px
                    else:
                        lx = self.left
                    d[size] = WndRect(lx, self.top, lx + px, self.bottom)
                self.area_mapping[pos] = d

        def set_vertical_rect(self) -> None:
            for pos in ("top", "bottom"):
                d = {}
                for size, px in self.possible_height.items():
                    if pos == "bottom":
                        ty = self.bottom - px
                    else:
                        ty = self.top
                    d[size] = WndRect(self.left, ty, self.right, ty + px)
                self.area_mapping[pos] = d

        @staticmethod
        def get_variation(max_size: int) -> dict:
            return {
                "small": int(max_size / 3),
                "middle": int(max_size / 2),
                "large": int(max_size * 2 / 3),
            }

    class CurrentMonitors:
        def __init__(self) -> None:
            self.monitors = []
            for mi in pyauto.Window.getMonitorInfo():
                mr = MonitorRect(rect=mi[1], is_primary=mi[2])
                mr.set_center_rect()
                mr.set_horizontal_rect()
                mr.set_vertical_rect()
                self.monitors.append(mr)

        def is_multi_widths(self) -> bool:
            variation = [m.max_width for m in self.monitors]
            return len(set(variation)) != 1

        def get_info(self) -> list:
            if self.is_multi_widths():
                return sorted(self.monitors, key=lambda x: x.max_width, reverse=True)
            return sorted(self.monitors, key=lambda x: x.is_primary, reverse=True)

    ################################
    # set cursor position
    ################################

    keymap_global["U0-Up"] = keymap.MouseWheelCommand(+0.5)
    keymap_global["U0-Down"] = keymap.MouseWheelCommand(-0.5)
    keymap_global["U0-Left"] = keymap.MouseHorizontalWheelCommand(-0.5)
    keymap_global["U0-Right"] = keymap.MouseHorizontalWheelCommand(+0.5)

    class CursorPos:
        def __init__(self) -> None:
            self._keymap = keymap
            self.pos = []
            monitors = CurrentMonitors().get_info()
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

    keymap_global["O-RCtrl"] = CursorPos().snap

    keymap_global["O-(236)"] = CursorPos().snap_to_center

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

        def __init__(self) -> None:
            self._keymap = keymap

        def alloc_flexible(self, km: Keymap) -> None:
            monitors = CurrentMonitors().get_info()
            for mod_mntr, mntr_idx in self.monitor_dict.items():
                for mod_area, size in self.size_dict.items():
                    for key, pos in self.snap_key_dict.items():
                        if mntr_idx < len(monitors):
                            wnd_rect = monitors[mntr_idx].area_mapping[pos][size]
                            km[mod_mntr + mod_area + key] = LazyFunc(wnd_rect.snap).defer()

        def alloc_maximize(self, km: Keymap, mapping_dict: dict) -> None:
            for key, towards in mapping_dict.items():

                def _snap() -> None:
                    VIRTUAL_FINGER.type_keys("LShift-LWin-" + towards)
                    delay()
                    self._keymap.getTopLevelWindow().maximize()

                km[key] = LazyFunc(_snap).defer()

    WndPosAllocator().alloc_flexible(keymap_global["U1-M"])

    WndPosAllocator().alloc_maximize(keymap_global, {"LC-U1-L": "Right", "LC-U1-H": "Left"})
    WndPosAllocator().alloc_maximize(
        keymap_global["U1-M"],
        {
            "U0-L": "Right",
            "U0-J": "Right",
            "U0-H": "Left",
            "U0-K": "Left",
        },
    )

    class WndShrinker:
        def __init__(self) -> None:
            self._keymap = keymap

        def invoke_snapper(self, horizontal: bool, default_pos: bool) -> Callable:
            def _snap() -> None:
                wnd = self._keymap.getTopLevelWindow()
                wr = WndRect(*wnd.getRect())
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

        def apply(self, km: Keymap) -> None:
            for key, params in {
                "H": {"horizontal": True, "default_pos": True},
                "L": {"horizontal": True, "default_pos": False},
                "K": {"horizontal": False, "default_pos": True},
                "J": {"horizontal": False, "default_pos": False},
            }.items():
                km["U1-" + key] = self.invoke_snapper(**params)

    WndShrinker().apply(keymap_global["U1-M"])

    ################################
    # pseudo espanso
    ################################

    def combo_mapper(root_map: Union[Callable, dict], keys: list, func: Callable) -> Union[Callable, dict, None]:
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
        def __init__(self) -> None:
            self._keymap = keymap
            self.mapping = self._keymap.defineMultiStrokeKeymap()
            user_path = UserPath()
            direct_puncher = KeyPuncher(defer_msec=50)
            for combo, stroke in {
                "X,X": (".txt"),
                "X,M": (".md"),
                "E,P,T": (":pekori: :tori_isogi: :houkoku_made:"),
                "E,P,G": (":pekori_gif:"),
                "P,M": (user_path.resolve(r"Sync\develop\app_setting\IME_google\convertion_dict\main.txt").path),
                "P,H": (user_path.resolve(r"Sync\develop\app_setting\IME_google\convertion_dict\human.txt").path),
                "P,D": (user_path.resolve(r"Desktop").path + "\\"),
                "M,P": ("=============================="),
                "M,D": ("div."),
                "M,S": ("span."),
                "N,0": ("0_plain"),
                "N,P,L": ("plain"),
                "N,P,A": ("proofed_by_author"),
                "N,P,J": ("project_proposal"),
                "N,P,P": ("proofed"),
                "N,S,A": ("send_to_author"),
                "N,S,P": ("send_to_printshop"),
                "Minus,H": ("\u2010"),
                "Minus,M": ("\u2014"),
                "Minus,N": ("\u2013"),
                "Minus,S": ("\u2212"),
                "Minus,D": ("\u30a0"),
                "R,B": (r"[\[［].+?[\]］]"),
                "R,P": (r"[\(（].+?[\)）]"),
                "C-R,S-F": ("(?!)", "Left"),
                "C-R,F": ("(?=)", "Left"),
                "C-R,S-P": ("(?<!)", "Left"),
                "C-R,P": ("(?<=)", "Left"),
                "A,E": ("\u00e9"),
                "A,S-E": ("\u00c9"),
                "C-A,E": ("\u00e8"),
                "C-A,S-E": ("\u00c8"),
                "U,S-A": ("\u00c4"),
                "U,A": ("\u00e4"),
                "U,S-O": ("\u00d6"),
                "U,O": ("\u00f6"),
                "U,S-U": ("\u00dc"),
                "U,U": ("\u00fc"),
            }.items():
                keys = combo.split(",")
                self.mapping = combo_mapper(self.mapping, keys, direct_puncher.invoke(*stroke))

            for i in "0123456789":
                keys = ("I", "S", i)
                s = chr(0x2080 + int(i))
                self.mapping = combo_mapper(self.mapping, keys, direct_puncher.invoke(s))

            indirect_puncher = KeyPuncher(recover_ime=True, defer_msec=50)
            for combo, stroke in {
                "Minus,F": ("\uff0d"),
                "F,G": ("\u3013\u3013"),
                "F,0": ("\u25cf\u25cf"),
                "F,4": ("\u25a0\u25a0"),
                "M,1": ("# "),
                "M,2": ("## "),
                "M,3": ("### "),
                "M,4": ("#### "),
                "M,5": ("##### "),
                "M,6": ("###### "),
            }.items():
                keys = combo.split(",")
                self.mapping = combo_mapper(self.mapping, keys, indirect_puncher.invoke(*stroke))

    keymap_global["U1-X"] = KeyCombo().mapping

    ################################
    # input customize
    ################################

    keymap_global["U0-Yen"] = KeyPuncher().invoke("S-Yen", "Left")
    keymap_global["U1-S"] = KeyPuncher().invoke("Slash")

    def punc_remap(mapping_dict: dict, km: Keymap, recover_ime: bool) -> None:
        for key, send in mapping_dict.items():
            km[key] = KeyPuncher(recover_ime=recover_ime).invoke(send)

    punc_remap(
        {
            "U0-1": "S-1",
            "U0-4": "$_",
            "U1-4": "$_.",
            "U0-Colon": "Colon",
            "U0-Comma": "Comma",
            "U0-Period": "Period",
            "U0-Slash": "Slash",
            "U0-U": "S-BackSlash",
            "U1-Enter": "<br />",
            "U1-Minus": "Minus",
            "U0-SemiColon": "SemiColon",
        },
        keymap_global,
        recover_ime=False,
    )

    punc_remap(
        {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Minus": "\u3000\u2015\u2015",
            "S-U0-Period": "\uff0e",  # FULLWIDTH FULL STOP
            "S-U0-U": "S-BackSlash",
            "U0-Minus": "\u2015\u2015",  # HORIZONTAL BAR * 2
            "U0-P": "\u30fb",  # KATAKANA MIDDLE DOT
            "S-C-U0-P": "\u2049",  # EXCLAMATION QUESTION MARK
            "S-U0-SemiColon": "+ ",
            "S-U0-8": "- ",
            "U1-G": "\u3013\u3013",  # GETA MARK * 2
            "LC-U1-B": "- ",
            "U1-1": "1. ",
            "S-U0-7": "1. ",
        },
        keymap_global,
        recover_ime=True,
    )

    class PairedPuncs:
        def __init__(self, pair_mapping: dict, recover_ime: bool) -> None:
            self.pair_mapping = pair_mapping
            self.recover_ime = recover_ime

        def _get_sender(self, pair: list) -> Callable:
            _, suffix = pair
            sent = pair + ["Left"] * len(suffix)
            return KeyPuncher(recover_ime=self.recover_ime).invoke(*sent)

        def apply_sender(self, km: Keymap) -> None:
            for key, pair in self.pair_mapping.items():
                km[key] = self._get_sender(pair)

        def _get_paster(self, pair: list) -> Callable:
            prefix, suffix = pair

            def _paster() -> None:
                IME_CONTROL.disable()
                ClipHandler().paste_current(lambda s: prefix + s + suffix)
                if self.recover_ime:
                    IME_CONTROL.enable()

            return _paster

        def apply_paster(self, km: Keymap, modifier: str = "LC-") -> None:
            for key, pair in self.pair_mapping.items():
                km[modifier + key] = self._get_paster(pair)

    PAIRS_WITHOUT_IME = PairedPuncs(
        {
            "U0-2": ['"', '"'],
            "U0-7": ["'", "'"],
            "U0-AtMark": ["`", "`"],
            "U1-AtMark": [" `", "` "],
            "U0-CloseBracket": ["[", "]"],
            "U1-9": ["(", ")"],
            "U1-CloseBracket": ["{", "}"],
            "U0-Caret": ["~~", "~~"],
        },
        False,
    )
    PAIRS_WITHOUT_IME.apply_sender(keymap_global)
    PAIRS_WITHOUT_IME.apply_paster(keymap_global, "LC-")

    PAIRS_WITH_IME = PairedPuncs(
        {
            "U0-8": ["\u300e", "\u300f"],  # WHITE CORNER BRACKET 『』
            "U0-9": ["\u3010", "\u3011"],  # BLACK LENTICULAR BRACKET 【】
            "U0-OpenBracket": ["\u300c", "\u300d"],  # CORNER BRACKET 「」
            "U0-Y": ["\u300a", "\u300b"],  # DOUBLE ANGLE BRACKET 《》
            "U1-2": ["\u201c", "\u201d"],  # DOUBLE QUOTATION MARK “”
            "U1-7": ["\u2018", "\u2019"],  # SINGLE QUOTATION MARK ‘’
            "U0-T": ["\u3014", "\u3015"],  # TORTOISE SHELL BRACKET 〔〕
            "U1-8": ["\uff08", "\uff09"],  # FULLWIDTH PARENTHESIS （）
            "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
            "U1-Y": ["\u3008", "\u3009"],  # ANGLE BRACKET 〈〉
        },
        True,
    )
    PAIRS_WITH_IME.apply_sender(keymap_global)
    PAIRS_WITH_IME.apply_paster(keymap_global, "LC-")

    class DirectInput:
        @staticmethod
        def invoke(key: str, turnoff_ime_later: bool = False) -> Callable:
            finish_keys = ["C-M"]
            if turnoff_ime_later:
                finish_keys.append("(243)")

            def _input() -> None:
                VIRTUAL_FINGER.type_keys(key)
                if IME_CONTROL.is_enabled():
                    VIRTUAL_FINGER.type_keys(*finish_keys)

            return _input

        @classmethod
        def invoke_puncs(cls, km: Keymap) -> None:
            for key, turnoff_ime in {
                "AtMark": True,
                "Caret": False,
                "CloseBracket": False,
                "Colon": False,
                "Comma": False,
                "LS-AtMark": True,
                "LS-Caret": False,
                "LS-Colon": True,
                "LS-Comma": True,
                "LS-Minus": False,
                "LS-Period": True,
                "LS-SemiColon": False,
                "LS-Slash": False,
                "LS-Yen": True,
                "OpenBracket": False,
                "Period": False,
                "SemiColon": False,
                "Slash": False,
                "Yen": True,
            }.items():
                km[key] = cls.invoke(key, turnoff_ime)

        @classmethod
        def invoke_alphabets(cls, km: Keymap) -> None:
            for alphabet in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                key = "LS-" + alphabet
                km[key] = cls.invoke(key, True)

        @classmethod
        def invoke_shift_numbers(cls, km: Keymap) -> None:
            for n in "123456789":
                key = "LS-" + n
                turnoff_ime = False
                if n in ("2", "3", "4"):
                    turnoff_ime = True
                km[key] = cls.invoke(key, turnoff_ime)

    DirectInput().invoke_puncs(keymap_global)
    DirectInput().invoke_alphabets(keymap_global)
    DirectInput().invoke_shift_numbers(keymap_global)

    keymap_global["BackSlash"] = DirectInput.invoke("S-BackSlash", False)
    keymap_global["C-U0-P"] = DirectInput.invoke("S-1", False)
    keymap_global["S-U0-P"] = DirectInput.invoke("S-Slash", False)

    class DateInput:
        @staticmethod
        def invoke(fmt: str, recover_ime: bool = False) -> Callable:
            def _input() -> None:
                d = datetime.datetime.today()
                IME_CONTROL.disable()
                seq = [c for c in d.strftime(fmt)]
                VIRTUAL_FINGER_QUICK.type_smart(*seq)
                if recover_ime:
                    IME_CONTROL.enable()

            return LazyFunc(_input).defer()

        @classmethod
        def apply(cls, km: Keymap) -> None:
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
                "J": ("%Y年%#m月%#d日", True),
            }.items():
                km[key] = cls.invoke(*params)

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

    class RemovableUnicode:
        alt = ord(" ")
        mapping = {}

        @classmethod
        def register(cls, ord: int) -> None:
            cls.mapping[ord] = cls.alt

        @classmethod
        def register_range(cls, pair: list) -> None:
            start, end = pair
            for i in range(int(start, 16), int(end, 16) + 1):
                cls.register(i)

        @classmethod
        def register_ranges(cls, pairs: list) -> None:
            for pair in pairs:
                cls.register_range(pair)

    class NoiseCleaner(RemovableUnicode):
        def __init__(self) -> None:
            super().__init__()
            self.register_ranges(
                [  # ascii
                    ["0021", "002F"],
                    ["003A", "0040"],
                    ["005B", "0060"],
                    ["007B", "007E"],
                ]
            )
            self.register_range(["2018", "201F"])  # quotation
            self.register_ranges(
                [  # horizontal bars
                    ["2010", "2017"],
                    ["2500", "2501"],
                    ["2E3A", "2E3B"],
                ]
            )
            self.register_ranges(
                [  # fullwidth symbols
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
            self.register_range(["2E80", "2EF3"])  # kangxi
            self.register(int("30FB", 16))  # KATAKANA MIDDLE DOT

        @classmethod
        def cleanup(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.mapping))

    NOISE_CLEANER = NoiseCleaner()

    class SearchQuery:
        def __init__(self, query: str) -> None:
            self._query = ""
            lines = query.strip().replace("\u200b", "").replace("\u3000", " ").replace("\t", " ").splitlines()
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
            for word in NOISE_CLEANER.cleanup(self._query).split(" "):
                if len(word):
                    if strict:
                        words.append('"{}"'.format(word))
                    else:
                        words.append(word)
            return urllib.parse.quote(" ".join(words))

    class WebSearcher:
        def __init__(self) -> None:
            self._keymap = keymap

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
                PathInfo(uri.format(query.encode(strict))).run()

            return LazyFunc(_search).defer()

        mod_dict = {
            "": (False, False),
            "S-": (True, False),
            "C-": (False, True),
            "S-C-": (True, True),
        }
        key_uri_dict = {
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

        def apply(self, km: Keymap) -> None:
            for mod_key, params in self.mod_dict.items():
                km[mod_key + "U0-S"] = self._keymap.defineMultiStrokeKeymap()
                for key, uri in self.key_uri_dict.items():
                    km[mod_key + "U0-S"][key] = self.invoke(uri, *params)

    WebSearcher().apply(keymap_global)

    ################################
    # activate window
    ################################

    class SystemBrowser:
        register_path = r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice"
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
        def __init__(self, exe_name: str, class_name: str) -> None:
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
            self.found = wnd.getLastActivePopup()
            return False

    class PseudoCuteExec:
        def __init__(self) -> None:
            self._keymap = keymap

        def activate_wnd(self, target: pyauto.Window) -> bool:
            if self._keymap.getWindow() == target:
                return False
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

        def invoke(self, exe_name: str, class_name: str, exe_path: str) -> Callable:
            scanner = WndScanner(exe_name, class_name)

            def _executer() -> None:
                scanner.scan()
                if scanner.found:
                    if not self.activate_wnd(scanner.found):
                        VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")
                else:
                    if exe_path:
                        PathInfo(exe_path).run()

            return LazyFunc(_executer).defer(80)

        activate_keymap = {
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
            "D": (
                "vivaldi.exe",
                "Chrome_WidgetWin_1",
                UserPath().resolve(r"AppData\Local\Vivaldi\Application\vivaldi.exe").path,
            ),
            "S": (
                "slack.exe",
                "Chrome_WidgetWin_1",
                UserPath().resolve(r"AppData\Local\slack\slack.exe").path,
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
                UserPath().resolve(r"scoop\apps\ksnip\current\ksnip.exe").path,
            ),
            "O": ("Obsidian.exe", "Chrome_WidgetWin_1", ""),
            "P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME", ""),
            "C-P": ("powerpnt.exe", "PPTFrameClass", ""),
            "E": ("EXCEL.EXE", "XLMAIN", ""),
            "I": ("i_view64.exe", "IrfanView", ""),
            "W": ("WINWORD.EXE", "OpusApp", ""),
            "V": ("Code.exe", "Chrome_WidgetWin_1", UserPath().get_editor()),
            "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1", ""),
            "T": ("TE64.exe", "TablacusExplorer", UserPath().get_filer()),
            "M": (
                "Mery.exe",
                "TChildForm",
                UserPath().resolve(r"AppData\Local\Programs\Mery\Mery.exe").path,
            ),
            "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
        }

        def apply_combo(self, km: Keymap) -> None:
            for key, params in self.activate_keymap.items():
                km[key] = self.invoke(*params)

        def apply_single(self, km: Keymap) -> None:
            for key, params in {
                "U1-T": ("TE64.exe", "TablacusExplorer", ""),
                "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME", ""),
                "LC-U1-M": (
                    "Mery.exe",
                    "TChildForm",
                    UserPath().resolve(r"AppData\Local\Programs\Mery\Mery.exe").path,
                ),
                "LC-U1-N": (
                    "notepad.exe",
                    "Notepad",
                    r"C:\Windows\System32\notepad.exe",
                ),
            }.items():
                km[key] = self.invoke(*params)

    PseudoCuteExec().apply_single(keymap_global)
    keymap_global["U1-C"] = keymap.defineMultiStrokeKeymap()
    keymap_global["U1-C"]["LC-M"] = lambda: PathInfo(r"C:\Personal\draft.txt").run()
    PseudoCuteExec().apply_combo(keymap_global["U1-C"])

    # invoke specific filer
    def invoke_filer(dir_path: str) -> Callable:
        filer_path = UserPath().get_filer()

        def _invoker() -> None:
            if PathInfo(dir_path).isAccessible:
                PathInfo(filer_path).run(dir_path)
            else:
                print("invalid-path: '{}'".format(dir_path))

        return LazyFunc(_invoker).defer()

    keymap_global["U1-F"] = keymap.defineMultiStrokeKeymap()
    keymap_global["U1-F"]["D"] = invoke_filer(UserPath().resolve(r"Desktop").path)
    keymap_global["U1-F"]["S"] = invoke_filer(r"X:\scan")

    def invoke_terminal() -> None:
        scanner = WndScanner("wezterm-gui.exe", "org.wezfurlong.wezterm")
        scanner.scan()
        if scanner.found:
            if not PseudoCuteExec().activate_wnd(scanner.found):
                VIRTUAL_FINGER.type_keys("C-AtMark")
        else:
            terminal_path = UserPath().resolve(r"scoop\apps\wezterm\current\wezterm-gui.exe").path
            PathInfo(terminal_path).run()

    keymap_global["LC-AtMark"] = LazyFunc(invoke_terminal).defer()

    def search_on_browser() -> None:
        if keymap.getWindow().getProcessName() == DEFAULT_BROWSER.get_exe_name():
            VIRTUAL_FINGER.type_keys("C-T")
        else:
            scanner = WndScanner(DEFAULT_BROWSER.get_exe_name(), DEFAULT_BROWSER.get_wnd_class())
            scanner.scan()
            if scanner.found:
                if PseudoCuteExec().activate_wnd(scanner.found):
                    delay()
                    VIRTUAL_FINGER.type_keys("C-T")
                else:
                    VIRTUAL_FINGER.type_keys("LCtrl-LAlt-Tab")
            else:
                PathInfo("https://duckduckgo.com").run()

    keymap_global["U0-Q"] = LazyFunc(search_on_browser).defer(100)

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=CheckWnd.is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"
    keymap_browser["O-LShift"] = "C-F"
    keymap_browser["LC-Q"] = "A-F4"

    def focus_chrome_content() -> None:
        wnd = keymap.getWindow()
        if wnd.getProcessName() == "chrome.exe":
            VIRTUAL_FINGER.type_keys("S-A-B", "F6")
        elif wnd.getProcessName() == "firefox.exe":
            VIRTUAL_FINGER.type_keys("C-L", "F6")

    keymap_browser["U0-F6"] = focus_chrome_content

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None

    # slack
    keymap_slack = keymap.defineWindowKeymap(exe_name="slack.exe", class_name="Chrome_WidgetWin_1")
    keymap_slack["O-LShift"] = "Esc", "C-K"
    keymap_slack["F3"] = "Esc", "C-K"
    keymap_slack["F1"] = KeyPuncher().invoke("+:")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")
    keymap_vscode["C-S-P"] = KeyPuncher().invoke("C-S-P")
    keymap_vscode["C-A-B"] = KeyPuncher().invoke("C-A-B")

    # mery
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")
    keymap_mery["LA-U0-J"] = "A-CloseBracket"
    keymap_mery["LA-U0-K"] = "A-OpenBracket"
    keymap_mery["LA-LC-U0-J"] = "A-C-CloseBracket"
    keymap_mery["LA-LC-U0-K"] = "A-C-OpenBracket"
    keymap_mery["LA-LS-U0-J"] = "A-S-CloseBracket"
    keymap_mery["LA-LS-U0-k"] = "A-S-OpenBracket"

    # cmder
    keymap_cmder = keymap.defineWindowKeymap(class_name="VirtualConsoleClass")
    keymap_cmder["LAlt-Space"] = "Lwin-LAlt-Space"
    keymap_cmder["C-W"] = lambda: None

    # sumatra PDF
    keymap_sumatra = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra)
    keymap_sumatra["O-LShift"] = KeyPuncher(recover_ime=True).invoke("F6", "C-Home", "C-F")
    keymap_sumatra["C-G"] = KeyPuncher().invoke("C-G")

    keymap_sumatra_inputmode = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra_inputmode)

    def sumatra_change_tab(km: Keymap) -> None:
        for key in ["C-Tab", "C-S-Tab"]:
            km[key] = "Esc", key

    sumatra_change_tab(keymap_sumatra_inputmode)

    keymap_sumatra_viewmode = keymap.defineWindowKeymap(check_func=CheckWnd.is_sumatra_viewmode)

    def sumatra_view_key(km: Keymap) -> None:
        for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            km[key] = KeyPuncher().invoke(key)

    sumatra_view_key(keymap_sumatra_viewmode)

    keymap_sumatra_viewmode["F"] = KeyPuncher(recover_ime=True).invoke("C-F")
    keymap_sumatra_viewmode["H"] = "C-S-Tab"
    keymap_sumatra_viewmode["L"] = "C-Tab"

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")
    keymap_word["F11"] = "A-F", "E", "P", "A"
    keymap_word["C-G"] = KeyPuncher().invoke("C-G")
    keymap_word["LC-Q"] = "A-F4"

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    keymap_ppt["F11"] = "A-F", "E", "P", "A"

    # excel
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")
    keymap_excel["F11"] = "A-F", "E", "P", "A"
    keymap_excel["LC-Q"] = "A-F4"

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
    keymap_tb["C-S-S"] = thunderbird_new_mail(["C-Home", "S-End", "C-X", "Delete", "A-S", "C-V"], ["A-S"])

    # filer
    keymap_filer = keymap.defineWindowKeymap(check_func=CheckWnd.is_filer_viewmode)
    KeyAllocator(
        {
            "A": ("Home"),
            "E": ("End"),
            "C": ("C-C"),
            "J": ("Down"),
            "K": ("Up"),
            "N": ("F2"),
            "R": ("C-R"),
            "U": ("LAlt-Up"),
            "V": ("C-V"),
            "W": ("C-W"),
            "X": ("C-X"),
            "Space": ("Enter"),
            "C-S-C": ("C-Add"),
            "C-L": ("A-D", "C-C"),
        }
    ).apply(keymap_filer)

    keymap_tablacus = keymap.defineWindowKeymap(check_func=CheckWnd.is_tablacus_viewmode)
    keymap_tablacus["H"] = "C-S-Tab"
    keymap_tablacus["L"] = "C-Tab"

    def tablacus_key_accelerator(km: Keymap) -> None:
        for alphabet in "ABCDEFGHIJKLMNOPQRSTUVWXTYZ":
            km[alphabet] = alphabet

    keymap_tablacus["LC-K"] = keymap.defineMultiStrokeKeymap()
    tablacus_key_accelerator(keymap_tablacus["LC-K"])

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
        @staticmethod
        def get_time(s) -> str:
            d_str = re.sub(r" 大阪.+$|^時[間刻]：", "", s)
            try:
                d = datetime.datetime.strptime(d_str, "%Y年%m月%d日 %I:%M %p")
                week = "月火水木金土日"[d.weekday()]
                return (d.strftime("%Y年%m月%d日（{}） %p %I:%M～")).format(week)
            except:
                return ""

        @classmethod
        def format(cls, copied: str) -> str:
            lines = copied.replace(": ", "\uff1a").strip().splitlines()
            if len(lines) < 9:
                print("ERROR: lack of lines.")
                return ""
            due = cls.get_time(lines[3])
            if len(due) < 1:
                print("ERROR: could not parse due date.")
                return ""
            return os.linesep.join(
                [
                    "------------------------------",
                    lines[2],
                    due,
                    lines[6],
                    lines[8],
                    lines[9],
                    "------------------------------",
                ]
            )

    class ClipboardMenu:
        def __init__(self) -> None:
            self._keymap = keymap

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
            if PathInfo(s).isAccessible:
                return Path(s).read_text("utf-8")
            return ""

        @staticmethod
        def skip_blank_line(s: str) -> str:
            lines = s.strip().splitlines()
            return os.linesep.join([l for l in lines if l.strip()])

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

        def get_menu_noise_reduction(self) -> list:
            return [
                (" Remove: - Blank lines ", self.format_cb(self.skip_blank_line)),
                ("         - Inside Paren ", self.replace_cb(r"[\uff08\u0028].+?[\uff09\u0029]", "")),
                ("         - Line-break ", self.replace_cb(r"\r?\n", "")),
                ("         - Quotations ", self.replace_cb(r"[\u0022\u0027]", "")),
                (" Fix: - Dumb Quotation ", self.format_cb(self.fix_dumb_quotation)),
                ("      - MSWord-Bullet ", self.replace_cb(r"\uf09f\u0009", "\u30fb")),
                ("      - KANGXI RADICALS ", self.format_cb(KangxiRadicals().fix)),
            ]

        def get_menu_transform(self) -> list:
            return [
                (" Transform: => A-Z/0-9 ", self.format_cb(CharWidth().to_half_letter)),
                ("            => \uff21-\uff3a/\uff10-\uff19 ", self.format_cb(CharWidth().to_full_letter)),
                ("            => abc ", lambda: ClipHandler.get_string().lower()),
                ("            => ABC ", lambda: ClipHandler.get_string().upper()),
                (" Comma: - Curly (\uff0c) ", self.replace_cb(r"\u3001", "\uff0c")),
                ("        - Straight (\u3001) ", self.replace_cb(r"\uff0c", "\u3001")),
            ]

        def get_menu_other(self) -> list:
            return [
                (" Cat local file ", self.format_cb(self.catanate_file_content)),
                (" Mask USERNAME ", self.format_cb(UserPath().mask_user_name)),
                (" Postalcode | Address ", self.format_cb(self.split_postalcode)),
                (" URL: - Decode ", self.format_cb(self.decode_url)),
                ("      - Shorten Amazon ", self.replace_cb(r"^.+amazon\.co\.jp/.+dp/(.{10}).*", r"https://www.amazon.jp/dp/\1")),
                (" Zoom invitation ", self.format_cb(Zoom().format)),
            ]

        def apply(self) -> None:
            for title, menu in {
                "Noise-Reduction": self.get_menu_noise_reduction(),
                "Transform Alphabet / Punctuation": self.get_menu_transform(),
                "Others": self.get_menu_other(),
            }.items():
                m = menu + [("---------------- EXIT ----------------", lambda: None)]
                self._keymap.cblisters += [(title, cblister_FixedPhrase(m))]

    ClipboardMenu().apply()
    keymap_global["LC-LS-X"] = LazyFunc(keymap.command_ClipboardList).defer(msec=100)
