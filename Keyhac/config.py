import datetime
import os
import fnmatch
import re
import time
import shutil
import subprocess
import urllib.parse
import unicodedata
import webbrowser
from enum import Enum
from typing import Union, Callable, Dict, List, Tuple, NamedTuple, TypeAlias
from pathlib import Path
from winreg import HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, OpenKey, QueryValueEx
from concurrent.futures import ThreadPoolExecutor

import ckit
import pyauto
import keyhac_ini
from keyhac import *
from keyhac_keymap import Keymap, KeyCondition, WindowKeymap
from keyhac_listwindow import ListWindow


def smart_check_path(path: Union[str, Path], timeout_sec: Union[int, float, None] = None) -> bool:
    """CASE-INSENSITIVE path check with timeout"""
    p = path if isinstance(path, Path) else Path(path)
    try:
        future = ThreadPoolExecutor(max_workers=1).submit(p.exists)
        return future.result(timeout_sec)
    except:
        return False


def check_fzf() -> bool:
    return shutil.which("fzf.exe") is not None


def open_vscode(*args: str) -> bool:
    try:
        if code_path := shutil.which("code"):
            cmd = [code_path] + list(args)
            subprocess.run(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
            return True
        return False
    except Exception as e:
        print(e)
        return False


def shell_exec(path: str, *args) -> None:
    if not isinstance(path, str):
        path = str(path)
    if path.startswith("http"):
        webbrowser.open(path)
        return
    path = os.path.expandvars(path)
    try:
        cmd = ["start", "", path] + list(args)
        subprocess.run(cmd, shell=True)
    except Exception as e:
        print(e)


def configure(keymap):

    def balloon(message: str) -> None:
        title = datetime.datetime.today().strftime("%Y%m%d-%H%M%S-%f")
        print(message)
        try:
            keymap.popBalloon(title, message, 1500)
        except:
            pass

    ################################
    # general setting
    ################################

    # console theme
    keymap.setFont("HackGen", 16)

    def set_custom_theme():
        name = "black"

        custom_theme = {
            "bg": "#3f3b39",
            "fg": "#a0b4a7",
            "cursor0": "#ffffff",
            "cursor1": "#ff4040",
            "bar_fg": "#000000",
            "bar_error_fg": "#ff4040",
            "select_bg": "#dff477",
            "select_fg": "#3f3b39",
            "caret0": "#ffffff",
            "caret1": "#ff0000",
        }
        ckit.ckit_theme.theme_name = name

        for k, v in custom_theme.items():
            rgb = tuple(int(v[i : i + 2], 16) for i in (1, 3, 5))
            ckit.ckit_theme.ini.set("COLOR", k, str(rgb))
        keymap.console_window.reloadTheme()

    set_custom_theme()

    # set console appearance
    keyhac_ini.setint("CONSOLE", "visible", 0)
    keyhac_ini.setint("CONSOLE", "x", 0)
    keyhac_ini.write()

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

    # keyboard macro
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["S-U0-0"] = keymap.command_RecordClear
    keymap_global["U1-0"] = keymap.command_RecordPlay
    keymap_global["U1-F4"] = keymap.command_RecordPlay
    keymap_global["C-U0-0"] = keymap.command_RecordPlay

    def bind_cursor_keys(wk: WindowKeymap) -> None:
        mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
        for mod_key in mod_keys:
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
                wk[mod_key + "U0-" + key] = mod_key + value

    bind_cursor_keys(keymap_global)

    def bind_keys(wk: WindowKeymap, bindig: dict) -> None:
        for key, value in bindig.items():
            wk[key] = value

    bind_keys(
        keymap_global,
        {
            # focus taskbar
            "U1-T": ("LWin-T"),
            # send n and space
            "LS-U0-N": ("N", "N", "Space"),
            # close
            "LC-Q": ("A-F4"),
            # delete around cursor
            "U0-Back": ("Back", "Delete"),
            # delete to bol / eol
            "S-U0-B": ("S-Home", "Delete"),
            "S-U0-D": ("S-End", "Delete"),
            # escape
            "O-(235)": ("Esc"),
            "U0-X": ("Esc"),
            # line selection
            "U1-A": ("End", "S-Home"),
            # punctuation
            "U0-4": ("S-4", "S-BackSlash"),
            "U0-Enter": ("Period"),
            "U0-U": ("S-BackSlash"),
            "U0-Z": ("Minus"),
            "U1-S": ("Slash"),
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
            # indent
            "U0-Tab": ("Home", "Tab"),
            # print
            "F1": ("C-P"),
            "U1-F1": ("F1"),
        },
    )

    def bind_paired_keys(wk: WindowKeymap, binding: dict) -> None:
        for key, value in binding.items():
            wk[key] = value, value, "Left"

    bind_paired_keys(
        keymap_global,
        {
            "U0-2": "LS-2",
            "U0-7": "LS-7",
            "U0-AtMark": "LS-AtMark",
        },
    )

    ################################
    # functions for custom hotkey
    ################################

    def delay(msec: int = 50) -> None:
        if 0 < msec:
            time.sleep(msec / 1000)

    PyautoInput: TypeAlias = Union[pyauto.Key, pyauto.KeyUp, pyauto.KeyDown, pyauto.Char]

    class Tap:
        mod: int = 0
        sequence: List[PyautoInput] = []

        def __init__(self, name: str):
            up = None
            tokens = [s for s in name.split("-")]

            for token in tokens[:-1]:
                t = token.strip().upper()
                try:
                    self.mod |= KeyCondition.strToMod(t, force_LR=True)
                except ValueError:
                    if up is not None:
                        continue
                    if t == "U":
                        up = True
                    else:
                        if t == "D":
                            up = False

            tail = tokens[-1]
            try:
                vk = KeyCondition.strToVk(tail.strip().upper())
                if up is None:
                    self.sequence = [pyauto.Key(vk)]
                else:
                    if up:
                        self.sequence = [pyauto.KeyUp(vk)]
                    else:
                        self.sequence = [pyauto.KeyDown(vk)]
            except ValueError:

                self.sequence = [pyauto.Char(c) for c in str(tail)]

    class VirtualFinger:
        def __init__(self, inter_stroke_pause: int = 10) -> None:
            self._inter_stroke_pause = inter_stroke_pause

        @staticmethod
        def begin() -> None:
            keymap.beginInput()
            keymap.setInput_Modifier(0)

        @staticmethod
        def end() -> None:
            keymap.endInput()

        @staticmethod
        def compile(*sequence: str) -> List[Tap]:
            return [Tap(elem) for elem in sequence]

        def send(self, *sequence: str) -> None:
            taps = self.compile(*sequence)
            self.send_compiled(taps)

        def send_compiled(self, taps: List[Tap]) -> None:
            for t in taps:
                delay(self._inter_stroke_pause)
                self.begin()
                keymap.setInput_Modifier(t.mod)
                for x in t.sequence:
                    keymap.input_seq.append(x)
                self.end()

    keymap.magical_key = VirtualFinger.compile("LWin-S-M", "U-Alt")
    keymap.mod_release_sequence = VirtualFinger.compile("U-Shift", "U-Alt", "U-Ctrl")

    def subthread_run(
        func: Callable,
        finished: Union[Callable, None] = None,
        focus_changed_in_subthread: bool = False,
    ) -> None:
        finger = VirtualFinger()
        if focus_changed_in_subthread:
            finger.send_compiled(keymap.magical_key)

        def _finished(job_item: ckit.JobItem) -> None:
            if finished is not None:
                finished(job_item)
            finger.send_compiled(keymap.mod_release_sequence)
            keymap.setInput_Modifier(0)

        job = ckit.JobItem(func, _finished)
        ckit.JobQueue.defaultQueue().enqueue(job)

    class SKKKey:
        toggle_vk = "(243)"
        kata = "Q"
        kana = "C-J"
        halfkata = "C-O"
        latin = "S-L"
        cancel = "Esc"
        reconv = "LWin-Slash"
        abbrev = "Slash"
        convpoint = "S-0"
        jlatin = "S-Q"
        prefix = "S-Period"

        @classmethod
        def taps(cls, *keys: str) -> List[Tap]:
            return VirtualFinger.compile(cls.kana, *keys)

    class ImeControl:
        def __init__(self, inter_stroke_pause: int = 10) -> None:
            self._finger = VirtualFinger(inter_stroke_pause)
            self.key_to_kana = SKKKey.taps()
            self.key_to_turnoff = SKKKey.taps(SKKKey.toggle_vk)
            self.key_to_kata = SKKKey.taps(SKKKey.kata)
            self.key_to_latin = SKKKey.taps(SKKKey.latin)
            self.key_to_abbrev = SKKKey.taps(SKKKey.abbrev)
            self.key_to_half_kata = SKKKey.taps(SKKKey.halfkata)
            self.key_to_full_latin = SKKKey.taps(SKKKey.jlatin)
            self.key_to_conv = SKKKey.taps(SKKKey.convpoint)
            self.key_to_reconv = SKKKey.taps(SKKKey.reconv, SKKKey.cancel)

        @staticmethod
        def get_status() -> int:
            return keymap.getWindow().getImeStatus()

        @staticmethod
        def set_status(mode: int) -> None:
            keymap.getWindow().setImeStatus(mode)
            delay(20)

        @classmethod
        def is_enabled(cls) -> bool:
            return cls.get_status() == 1

        @classmethod
        def enable(cls) -> None:
            if not cls.is_enabled():
                cls.set_status(1)

        @classmethod
        def disable(cls) -> None:
            if cls.is_enabled():
                cls.set_status(0)

        def turnoff_skk(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_turnoff)

        def to_skk_kana(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_kana)

        def to_skk_latin(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_latin)

        def to_skk_abbrev(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_abbrev)

        def to_skk_kata(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_kata)

        def to_skk_half_kata(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_half_kata)

        def to_skk_full_latin(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_full_latin)

        def start_skk_conv(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_conv)

        def reconvert_with_skk(self) -> None:
            self.enable()
            self._finger.send_compiled(self.key_to_reconv)

    def apply_ime_control() -> None:
        control = ImeControl()
        for key, func in {
            "U1-J": control.to_skk_kana,
            "LC-U0-I": control.to_skk_kata,
            "U0-F7": control.to_skk_kata,
            "U0-O": control.to_skk_half_kata,
            "LC-LS-U0-I": control.to_skk_half_kata,
            "U0-F8": control.to_skk_half_kata,
            "U0-F": control.disable,
            "LS-U0-F": control.to_skk_kana,
            "S-U1-J": control.to_skk_latin,
            "U1-I": control.reconvert_with_skk,
            "O-(236)": control.to_skk_abbrev,
            "LS-(236)": control.start_skk_conv,
        }.items():
            keymap_global[key] = func

    apply_ime_control()

    def suppress_binded_key(func: Callable, wait_msec: int = 20) -> Callable:
        def _calmdown(_) -> None:
            delay(wait_msec)

        def _wrapper() -> None:
            keymap.setInput_Modifier(0)
            subthread_run(_calmdown, lambda _: func())

        return _wrapper

    class ClipHandler:
        @staticmethod
        def get_string() -> str:
            try:
                return ckit.getClipboardText() or ""
            except:
                return ""

        @staticmethod
        def set_string(s: str) -> None:
            try:
                ckit.setClipboardText(str(s))
            except:
                pass

        @staticmethod
        def send_copy_key():
            VirtualFinger().send("C-C")

        @staticmethod
        def send_paste_key():
            VirtualFinger().send("C-V")

        @classmethod
        def paste(
            cls, s: Union[str, None] = None, format_func: Union[Callable, None] = None
        ) -> None:
            if s is None:
                s = cls.get_string()
                if len(s) < 1:
                    # empty clipboard text may means image inside clipboard.
                    cls.send_paste_key()
                    return
            if format_func is not None:
                s = format_func(s)
            cls.set_string(s)
            cls.send_paste_key()

        def after_copy(cls, deferred: Callable) -> None:
            cb = cls.get_string()
            cls.send_copy_key()

            def _watch_clipboard(job_item: ckit.JobItem) -> None:
                job_item.origin = cb
                job_item.copied = ""
                trial = 600
                delay(120)
                for _ in range(trial):
                    try:
                        s = cls.get_string()
                        if not s.strip():
                            continue
                        if s != job_item.origin:
                            job_item.copied = s
                            break
                    except Exception as e:
                        print(e)
                        break

            subthread_run(_watch_clipboard, deferred)

    keymap_global["U0-V"] = ClipHandler().paste

    class FIFOStack:
        def __init__(self) -> None:
            self.items = []
            self.enabled = False

        def _enable(self) -> None:
            balloon("FIFO mode ON!")
            self.enabled = True

        def _disable(self) -> None:
            balloon("FIFO mode OFF!")
            self.enabled = False

        def toggle(self) -> None:
            if self.enabled:
                self._disable()
            else:
                self._enable()

        def register(self, s: str) -> None:
            if self.enabled:
                self.items.append(s)
                msg = "FIFO stack total: {}".format(self.count())
                balloon(msg)
            else:
                balloon("FIFO mode is not enabled.")

        def reset(self) -> None:
            self.items = []

        def bulk_register(self, lines: str) -> None:
            if self.enabled:
                self.reset()
                self.items = [line for line in lines.splitlines() if line.strip()]
                msg = "FIFO stack total: {}".format(self.count())
                balloon(msg)
            else:
                balloon("FIFO mode is not enabled.")

        def count(self) -> int:
            return len(self.items)

        def pop(self) -> Union[str, None]:
            if not self.enabled:
                balloon("FIFO mode is not enabled.")
                return None
            if 0 < self.count():
                cb = self.items.pop(0)
                balloon("FIFO stack rest: {}".format(self.count()))
                if self.count() < 1:
                    self._disable()
                return cb
            return None

    keymap.fifo_stack = FIFOStack()

    keymap_global["LC-LS-U0-X"] = keymap.fifo_stack.toggle

    keymap_global["LC-LS-U0-F"] = lambda: keymap.fifo_stack.bulk_register(ClipHandler.get_string())

    def smart_copy():
        if keymap.fifo_stack.enabled:

            def _register(job_item):
                cb = job_item.copied
                if cb:
                    keymap.fifo_stack.register(cb)

            ClipHandler().after_copy(_register)
        else:
            ClipHandler().send_copy_key()

    keymap_global["LC-C"] = smart_copy

    def smart_paste() -> None:
        if 0 < keymap.fifo_stack.count() and keymap.fifo_stack.enabled:
            cb = keymap.fifo_stack.pop()
            ClipHandler().paste(cb)
        else:
            ClipHandler().send_paste_key()

    keymap_global["LC-V"] = smart_paste

    ################################
    # custom hotkey
    ################################

    class StrCleaner:
        @staticmethod
        def remove_whitespace(s: str) -> str:
            return s.strip().translate(
                str.maketrans(
                    "",
                    "",
                    "\u0009\u0020\u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c\u200d\u200e\u200f\u202f\u205f\u3000\ufeff",
                )
            )

        @classmethod
        def invoke_paster(cls, no_space: bool = False, no_break: bool = False) -> Callable:
            def _clean(s) -> str:
                s = s.strip()
                if no_space:
                    s = cls.remove_whitespace(s)
                if no_break:
                    s = "".join(s.splitlines())
                return s

            def _paste() -> None:
                ClipHandler().paste(format_func=_clean)

            return _paste

        @classmethod
        def apply(cls, km: WindowKeymap, key: str) -> None:
            for mod1, no_space in {
                "": False,
                "C-": True,
            }.items():
                for mod2, no_break in {
                    "": False,
                    "S-": True,
                }.items():
                    km[mod1 + mod2 + key] = cls.invoke_paster(no_space, no_break)

    keymap_global["U1-V"] = keymap.defineMultiStrokeKeymap()
    StrCleaner().apply(keymap_global["U1-V"], "V")

    # paste with quote mark
    class Quoter:
        @staticmethod
        def simple_quote(s: str) -> str:
            lines = s.strip().splitlines()
            return "\n".join([keymap.quote_mark + line for line in lines])

        @staticmethod
        def as_single_line(s: str) -> str:
            lines = s.strip().splitlines()
            return keymap.quote_mark + "".join([line.strip() for line in lines])

        @staticmethod
        def skip_blank_line(s: str) -> str:
            lines = []
            for line in s.strip().splitlines():
                if 0 < len(line.strip()):
                    lines.append(keymap.quote_mark + line)
                else:
                    lines.append("")
            return "\n".join(lines)

        @staticmethod
        def invoke_paster(func: Callable) -> Callable:
            def _paster() -> None:
                ClipHandler().paste(None, func)

            return _paster

        @classmethod
        def apply(cls, km: WindowKeymap, key: str) -> None:
            km[key] = cls.invoke_paster(cls.simple_quote)
            km["C-" + key] = cls.invoke_paster(cls.as_single_line)
            km["S-" + key] = cls.invoke_paster(cls.skip_blank_line)

    Quoter().apply(keymap_global, "U1-Q")

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
    # config keys
    ################################

    def reload_config() -> None:
        ckit.JobQueue.cancelAll()
        keymap.configure()
        keymap.updateKeymap()
        ts = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
        balloon("{} reloaded config.py".format(ts))

    keymap_global["U1-F12"] = suppress_binded_key(reload_config, 60)

    def open_keyhac_repo() -> None:
        config_path = os.path.join(os.environ.get("APPDATA"), "Keyhac")
        if not smart_check_path(config_path):
            balloon("config not found: {}".format(config_path))
            return

        dir_path = config_path
        if (real_path := os.path.realpath(config_path)) != dir_path:
            dir_path = os.path.dirname(real_path)

        def _open(_) -> None:
            result = open_vscode(dir_path)
            if not result:
                shell_exec(dir_path)

        subthread_run(_open)

    keymap.editor = lambda _: open_keyhac_repo()

    keymap_global["U0-F12"] = open_keyhac_repo

    # clipboard menu
    def clipboard_history_menu() -> None:
        def _menu(_) -> None:
            keymap.command_ClipboardList()

        subthread_run(_menu)

    keymap_global["LC-LS-X"] = suppress_binded_key(clipboard_history_menu, 60)

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
    keymap_global["U1-M"]["X"] = lambda: keymap.getTopLevelWindow().maximize()
    keymap_global["U1-M"]["N"] = lambda: keymap.getTopLevelWindow().minimize()

    class RectEdge(Enum):
        left = 0
        top = 1
        right = 2
        bottom = 3

    class Rect:
        min_width = 300
        min_height = 200

        def __init__(self, left: int, top: int, right: int, bottom: int) -> None:
            self.left = left
            self.top = top
            self.right = right
            self.bottom = bottom
            self.width = right - left
            self.height = bottom - top

        def move_edge(self, toward: RectEdge, delta: int) -> List[int]:
            r = [
                self.left,
                self.top,
                self.right,
                self.bottom,
            ]
            opposite = (toward.value + 2) % 4
            r[opposite] = r[toward.value] + delta
            return r

        def resize(self, scale: float, toward: RectEdge) -> List[int]:
            if toward in [RectEdge.left, RectEdge.right]:
                dim = self.width
            else:
                dim = self.height
            delta = int(dim * scale)
            if toward in [RectEdge.right, RectEdge.bottom]:
                delta = delta * -1
            return self.move_edge(toward, delta)

    def apply_window_snapper(km: WindowKeymap) -> None:
        altkey_status = [False, True]
        scale_mapping = {
            "": 1 / 2,
            "S-": 2 / 3,
            "C-": 1 / 3,
        }
        edge_mapping = {
            "H": RectEdge.left,
            "L": RectEdge.right,
            "J": RectEdge.bottom,
            "K": RectEdge.top,
        }

        for alt_pressed in altkey_status:
            for area_mod, scale in scale_mapping.items():
                for key, edge in edge_mapping.items():

                    def _invoke(
                        a: bool = alt_pressed, s: float = scale, e: RectEdge = edge
                    ) -> Callable:

                        def __get_new_rect() -> List[int]:
                            infos = pyauto.Window.getMonitorInfo()
                            infos.sort(key=lambda info: info[2] != 1)
                            target = infos[1] if 1 < len(infos) and a else infos[0]
                            monitor_work_rect = Rect(*target[1])
                            return monitor_work_rect.resize(s, e)

                        def __snap() -> None:
                            wnd = keymap.getTopLevelWindow()
                            if not wnd or CheckWnd.is_keyhac_console(wnd):
                                return
                            rect = __get_new_rect()
                            if wnd.getRect() == rect:
                                return

                            def _job_snap(_) -> None:
                                if wnd.isMaximized():
                                    wnd.restore()
                                    delay()
                                wnd.setRect(rect)

                            def _job_finished(_) -> None:
                                if wnd.getRect() != rect:
                                    wnd.setRect(rect)

                            subthread_run(_job_snap, _job_finished)

                        return __snap

                    mod_alt = "A-" if alt_pressed else ""
                    km[mod_alt + area_mod + key] = _invoke()

    apply_window_snapper(keymap_global["U1-M"])

    def apply_maximized_window_snapper(km: WindowKeymap, keybinding: dict) -> None:
        finger = VirtualFinger()
        for key, towards in keybinding.items():

            def _snap() -> None:
                def __maximize(_) -> None:
                    keymap.getTopLevelWindow().maximize()

                def __snapper(_) -> None:
                    finger.send("LShift-LWin-" + towards)

                subthread_run(__maximize, __snapper)

            km[key] = _snap

    apply_maximized_window_snapper(keymap_global, {"LC-U1-L": "Right", "LC-U1-H": "Left"})
    apply_maximized_window_snapper(
        keymap_global["U1-M"],
        {
            "U0-L": "Right",
            "U0-J": "Right",
            "U0-H": "Left",
            "U0-K": "Left",
        },
    )

    class WndShrinker:

        @staticmethod
        def invoke_snapper(toward: RectEdge) -> Callable:

            def _snapper() -> None:
                def __snap(_) -> None:
                    wnd = keymap.getTopLevelWindow()
                    rect = wnd.getRect()
                    resized = Rect(*rect).resize(0.5, toward)
                    if wnd.isMaximized():
                        wnd.restore()
                        delay()
                    wnd.setRect(resized)

                subthread_run(__snap)

            return _snapper

        @classmethod
        def apply(cls, km: WindowKeymap) -> None:
            for key, toward in {
                "H": RectEdge.left,
                "L": RectEdge.right,
                "K": RectEdge.top,
                "J": RectEdge.bottom,
            }.items():
                km["U1-" + key] = cls.invoke_snapper(toward)

    WndShrinker().apply(keymap_global["U1-M"])

    class MonitorCenterAvoider:

        @staticmethod
        def get_border_x(wnd: pyauto.Window) -> int:
            x = -1
            rect = Rect(*wnd.getRect())
            center_x = int((rect.left + rect.right) / 2)
            monitors = pyauto.Window.getMonitorInfo()
            for monitor in monitors:
                monitor_rect = Rect(*monitor[1])
                if monitor_rect.left <= center_x and center_x <= monitor_rect.right:
                    x = int((monitor_rect.left + monitor_rect.right) / 2)
                    break
            return x

        @classmethod
        def invoke_avoider(cls, show_left: bool) -> Callable:

            def _avoider() -> None:
                def _snap(_) -> None:
                    wnd = keymap.getTopLevelWindow()
                    if CheckWnd.is_keyhac_console(wnd) or wnd.isMaximized():
                        return
                    border = cls.get_border_x(wnd)
                    if border < 0:
                        return
                    rect = list(wnd.getRect())
                    width = Rect(*rect).width
                    if rect[0] == border:
                        rect[0] = border - width
                        rect[2] = rect[0] + width
                    elif rect[2] == border:
                        rect[0] = border
                        rect[2] = rect[0] + width
                    else:
                        i = 0 if show_left else 2
                        rect[i] = border
                    wnd.setRect(rect)

                subthread_run(_snap)

            return _avoider

    keymap_global["U1-M"]["OpenBracket"] = MonitorCenterAvoider().invoke_avoider(False)
    keymap_global["U1-M"]["CloseBracket"] = MonitorCenterAvoider().invoke_avoider(True)

    ################################
    # set cursor position
    ################################

    class CursorPos:
        @staticmethod
        def get_pos() -> list:
            infos = pyauto.Window.getMonitorInfo()
            infos.sort(key=lambda info: info[2] != 1)
            rects = [Rect(*info[1]) for info in infos]
            pos = []
            for rect in rects:
                for i in (1, 3):
                    y = rect.top + int(rect.height / 2)
                    x = rect.left + int(rect.width / 4) * i
                    pos.append([x, y])
            return pos

        @classmethod
        def snap(cls) -> None:
            pos = cls.get_pos()
            x, y = pyauto.Input.getCursorPos()
            idx = -1
            for i, p in enumerate(pos):
                if p[0] == x and p[1] == y:
                    idx = i
            if idx < 0 or idx == len(pos) - 1:
                cls.set_position(*pos[0])
            else:
                cls.set_position(*pos[idx + 1])

        @staticmethod
        def set_position(x: int, y: int) -> None:
            keymap.beginInput()
            keymap.input_seq.append(pyauto.MouseMove(x, y))
            keymap.endInput()

        @classmethod
        def snap_to_center(cls) -> None:
            wnd = keymap.getTopLevelWindow()
            wnd_left, wnd_top, wnd_right, wnd_bottom = wnd.getRect()
            to_x = int((wnd_left + wnd_right) / 2)
            to_y = int((wnd_bottom + wnd_top) / 2)
            cls.set_position(to_x, to_y)

    keymap_global["O-RCtrl"] = CursorPos().snap
    keymap_global["O-RShift"] = CursorPos().snap_to_center

    ################################
    # input customize
    ################################

    class SKKSender:
        def __init__(
            self,
            inter_stroke_pause: int = 0,
        ) -> None:
            self.finger = VirtualFinger(inter_stroke_pause)
            self.control = ImeControl(inter_stroke_pause)

        def invoke(self, mode_setter: Callable, *sequence) -> Callable:
            taps = self.finger.compile(*sequence)

            def _send() -> None:
                mode_setter()
                self.finger.send_compiled(taps)

            return _send

        def under_kanamode(self, *sequence) -> Callable:
            sender = self.invoke(self.control.to_skk_kana, *sequence)
            return sender

        def under_latinmode(self, *sequence) -> Callable:
            sender = self.invoke(self.control.to_skk_latin, *sequence)
            return sender

        def without_mode(self, *sequence) -> Callable:
            sender = self.invoke(self.control.disable, *sequence)
            return sender

    # select-to-left with ime control
    keymap_global["U1-B"] = SKKSender().under_kanamode("S-Left")
    keymap_global["LS-U1-B"] = SKKSender().under_kanamode("S-Right")
    keymap_global["U1-Space"] = SKKSender().under_kanamode("C-S-Left")
    keymap_global["U1-N"] = SKKSender().under_kanamode("C-S-Left", SKKKey.convpoint, "S-4", "Tab")
    keymap_global["U1-4"] = SKKSender().under_kanamode(SKKKey.convpoint, "S-4")

    class DirectSender:
        def __init__(self, recover_ime: bool = True) -> None:
            self.skk = SKKSender(inter_stroke_pause=0)
            self.recover = recover_ime

        def invoke(self, *sequence) -> Callable:
            # use this function instead of `self.control.disable()` in order to confirm buffers by Ctrl-J while conversion.
            def _turnoff() -> None:
                if self.skk.control.is_enabled():
                    self.skk.control.turnoff_skk()

            seq = list(sequence)
            if self.recover:
                seq.append(SKKKey.toggle_vk)
            return self.skk.invoke(_turnoff, *seq)

        def bind(self, km: WindowKeymap, binding: dict) -> None:
            for key, sent in binding.items():
                km[key] = self.invoke(sent)

        def bind_circumfix(self, km: WindowKeymap, binding: dict) -> None:
            for key, circumfix in binding.items():
                _, suffix = circumfix
                sequence = circumfix + ["Left"] * len(suffix)
                km[key] = self.invoke(*sequence)

    keymap_global["U0-M"] = keymap.defineMultiStrokeKeymap()

    def replace_last_nchar(km: WindowKeymap, newstr: str) -> Callable:
        for n in "0123":
            seq = ["Back"] * int(n) + [newstr, SKKKey.kana]
            km[n] = SKKSender().under_latinmode(*seq)

    replace_last_nchar(keymap_global["U0-M"], "先生")

    # markdown list
    keymap_global["S-U0-8"] = DirectSender().invoke("U-Shift", "Minus", " ")
    keymap_global["U1-1"] = DirectSender().invoke("1.", " ")

    DirectSender().bind(
        keymap_global,
        {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Minus": "\u3000\u2015\u2015",
            "S-U0-Period": "\uff0e",  # FULLWIDTH FULL STOP
            "U0-Minus": "\u2015\u2015",  # HORIZONTAL BAR * 2
        },
    )

    DirectSender().bind_circumfix(
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
            "U1-Y": ["\u300a", "\u300b"],  # DOUBLE ANGLE BRACKET
            "U1-K": ["\u3014", "\u3015"],  # TORTOISE BRACKET
            "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
        },
    )

    DirectSender(False).bind(
        keymap_global,
        {
            "U0-1": "S-1",
            "U0-Colon": "Colon",
            "U0-Slash": "Slash",
            "U1-Minus": "Minus",
            "U0-SemiColon": "SemiColon",
            "U1-SemiColon": "+:",
            "U1-X": "!",
            "LS-U1-X": "!=",
            "U0-Comma": "Comma",
            "U0-Period": "Period",
        },
    )
    DirectSender(False).bind_circumfix(
        keymap_global,
        {
            "U0-CloseBracket": ["[", "]"],
            "U1-9": ["(", ")"],
            "U1-CloseBracket": ["{", "}"],
            "U0-Caret": ["~~", "~~"],
            "U1-Comma": ["<", ">"],
            "LS-U1-Comma": ["</", ">"],
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
            self.mapping = {}

        def register(self, charcode: int) -> None:
            self.mapping[charcode] = self._repl

        def register_range(self, start: int, end: int) -> None:
            for i in range(start, end + 1):
                self.register(i)

    class SearchNoise(UnicodeMapper):
        def __init__(self):
            super().__init__(" ")
            self.register(0x30FB)  # KATAKANA MIDDLE DOT
            self.register_range(0x2018, 0x201F)  # quotation
            self.register_range(0x2E80, 0x2EF3)  # kangxi
            noises = (
                [
                    # ascii
                    (0x0021, 0x002F),
                    (0x003A, 0x0040),
                    (0x005B, 0x0060),
                    (0x007B, 0x007E),
                ]
                + [
                    # bars
                    (0x2010, 0x2017),
                    (0x2500, 0x2501),
                    (0x2E3A, 0x2E3B),
                ]
                + [
                    # fullwidth
                    (0x25A0, 0x25EF),
                    (0x3000, 0x3004),
                    (0x3008, 0x3040),
                    (0x3097, 0x30A0),
                    (0x3097, 0x30A0),
                    (0x30FD, 0x30FF),
                    (0xFF01, 0xFF0F),
                    (0xFF1A, 0xFF20),
                    (0xFF3B, 0xFF40),
                    (0xFF5B, 0xFF65),
                ]
            )
            for noise in noises:
                self.register_range(*noise)

        def cleanup(self, s: str) -> str:
            return s.translate(str.maketrans(self.mapping))

    class SearchQuery:
        noise_mapping = SearchNoise()

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
            for honor in ["監修", "共著", "共編著", "編著", "共編", "分担執筆", "et al."]:
                self._query = self._query.replace(honor, " ")

        def remove_hiragana(self) -> None:
            self._query = re.sub(r"[\u3041-\u3093]", " ", self._query)

        def encode(self, strict: bool = False) -> str:
            words = []
            for word in self.noise_mapping.cleanup(self._query).split(" "):
                if len(word):
                    if strict:
                        words.append('"{}"'.format(word))
                    else:
                        words.append(word)
            return urllib.parse.quote(" ".join(words))

    class WebSearcher:
        def __init__(self, uri_mapping: dict) -> None:
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

            return _searcher

        def apply(self, km: WindowKeymap) -> None:
            for shift_key in ("", "S-"):
                for ctrl_key in ("", "C-"):
                    is_strict = 0 < len(shift_key)
                    strip_hiragana = 0 < len(ctrl_key)
                    trigger_key = shift_key + ctrl_key + "U0-S"
                    km[trigger_key] = keymap.defineMultiStrokeKeymap()
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
            "M": "https://www.merriam-webster.com/dictionary/{}",
            "N": "https://ndlsearch.ndl.go.jp/search?cs=bib&f-ht=ndl&keyword={}",
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
            ext = ".exe"
            c = cls.get_commandline()
            return c[: c.find(ext) + len(ext)].strip('"')

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

    keymap.default_browser = SystemBrowser()

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
                except Exception as e:
                    print("Failed to activate window due to exception:", e)
                    return False

            print("Failed to activate window due to timeout.")
            return False

    class PseudoCuteExec:
        @staticmethod
        def invoke(exe_name: str, class_name: str = "", exe_path: str = "") -> Callable:
            def _executor() -> None:

                def _activate(job_item: ckit.JobItem) -> None:
                    job_item.result = None
                    scanner = WndScanner(exe_name, class_name)
                    scanner.scan()
                    wnd = scanner.found
                    if wnd is None:
                        if exe_path:
                            if isinstance(exe_path, Callable):
                                exe_path()
                            else:
                                shell_exec(exe_path)
                    else:
                        job_item.result = WindowActivator(wnd).activate()

                def _finished(job_item: ckit.JobItem) -> None:
                    if job_item.result is not None:
                        if not job_item.result:
                            VirtualFinger().send("LCtrl-LAlt-Tab")

                subthread_run(_activate, _finished, True)

            return _executor

        @classmethod
        def apply(cls, wnd_keymap: WindowKeymap, remap_table: dict = {}) -> None:
            for key, params in remap_table.items():
                func = cls.invoke(*params)
                wnd_keymap[key] = suppress_binded_key(func, 120)

    PseudoCuteExec().apply(
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
    PseudoCuteExec().apply(
        keymap_global["U1-C"],
        {
            "Space": (
                keymap.default_browser.get_exe_name(),
                keymap.default_browser.get_wnd_class(),
                keymap.default_browser.get_exe_path(),
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
            "V": ("Code.exe", "Chrome_WidgetWin_1", open_vscode),
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
                if CheckWnd.is_keyhac_console(wnd):
                    return True
                if wnd.getProcessName() in ignore_list:
                    return True
                if not wnd.getText():
                    return True
                if popup := wnd.getLastActivePopup():
                    n = popup.getProcessName().replace(".exe", "")
                    if t := popup.getText():
                        n += "[{}]".format(t)
                    popup_table[n] = popup
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
            wnd = popup_table.get(result, None)
            if wnd is not None:
                job_item.result = WindowActivator(wnd).activate()

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.result is not None:
                if not job_item.result:
                    VirtualFinger().send("LCtrl-LAlt-Tab")

        subthread_run(_fzf_wnd, _finished, True)

    keymap_global["U1-E"] = suppress_binded_key(fuzzy_window_switcher)

    def invoke_draft() -> None:
        def _invoke(_) -> None:
            shell_exec(r"${USERPROFILE}\Personal\draft.txt")

        subthread_run(_invoke)

    keymap_global["LS-LC-U1-M"] = invoke_draft

    def search_on_browser() -> None:
        finger = VirtualFinger(20)
        if keymap.getWindow().getProcessName() == keymap.default_browser.get_exe_name():
            finger.send("C-T")
            return

        def _activate(job_item: ckit.JobItem) -> None:
            delay()
            job_item.result = None
            scanner = WndScanner(
                keymap.default_browser.get_exe_name(), keymap.default_browser.get_wnd_class()
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

        subthread_run(_activate, _finished, True)

    keymap_global["U0-Q"] = search_on_browser

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=CheckWnd.is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None

    # rsturio
    keymap_rstudio = keymap.defineWindowKeymap(exe_name="rstudio.exe")
    keymap_rstudio["U0-Minus"] = DirectSender(False).invoke("S-Comma", "Minus")

    # slack
    keymap_slack = keymap.defineWindowKeymap(exe_name="slack.exe", class_name="Chrome_WidgetWin_1")
    keymap_slack["F3"] = DirectSender(False).invoke("C-K")
    keymap_slack["C-K"] = keymap_slack["F3"]
    keymap_slack["C-E"] = keymap_slack["F3"]
    keymap_slack["F1"] = DirectSender(False).invoke("S-SemiColon", "Colon")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")

    def remap_vscode(*keys: str) -> Callable:
        sender = DirectSender(False)
        for key in keys:
            keymap_vscode[key] = sender.invoke(key)

    remap_vscode("C-E", "C-F", "C-S-F", "C-S-E", "C-S-G", "RC-RS-X", "C-0", "C-S-P", "C-A-B")

    # mery
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")

    def remap_mery(binding: dict) -> Callable:
        for key, value in binding.items():
            keymap_mery[key] = value

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
        }
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

    def sumatra_view_key() -> None:
        sender = DirectSender(False)
        for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            keymap_sumatra_viewmode[key] = sender.invoke(key)

    sumatra_view_key()

    keymap_sumatra_viewmode["H"] = "C-S-Tab"
    keymap_sumatra_viewmode["L"] = "C-Tab"

    def office_to_pdf(km: WindowKeymap, key: str = "F11") -> None:
        km[key] = "A-F", "E", "P", "A"

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")
    keymap_word["S-C-R"] = "C-H"
    keymap_word["RC-R"] = keymap_word["S-C-R"]
    keymap_word["RC-H"] = keymap_word["S-C-R"]
    office_to_pdf(keymap_word)

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    office_to_pdf(keymap_ppt)
    keymap_ppt["O-(236)"] = ImeControl(40).to_skk_abbrev

    # excel
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")
    office_to_pdf(keymap_excel)

    def select_all() -> None:
        finger = VirtualFinger()
        if keymap.getWindow().getClassName() == "EXCEL6":
            finger.send("C-End", "C-S-Home")
        else:
            finger.send("C-A")

    keymap_excel["C-A"] = select_all

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
                "参加手順",
                "invitations?signature=",
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

    class NestedCircumfix:
        def __init__(self, prime_pair: tuple, secondary_pair: tuple):
            self.pairs = [prime_pair, secondary_pair]

        def fix(self, s: str) -> str:
            stack = []
            result = list(s)
            openChar, closeChar = self.pairs[0]
            for i, char in enumerate(s):
                if char == openChar:
                    stack.append(i)
                elif char == closeChar:
                    if stack:
                        start = stack.pop()
                        depth = len(stack)
                        left, right = self.pairs[depth % 2]
                        result[start] = left
                        result[i] = right

            return "".join(result)

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
        def to_list(s: str) -> str:
            lines = s.splitlines()
            return "\n".join(["- " + l for l in lines])

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
            reg = re.compile(r"先生$|様$|(先生|様)(?=[、。：；（）［］・！？\s])")
            return reg.sub("", s)

        @staticmethod
        def trim_space_on_line_head(s: str) -> str:
            return "\n".join([line.lstrip() for line in s.splitlines()])

        @staticmethod
        def format_nested_paren(s: str) -> str:
            return NestedCircumfix(("（", "）"), ("〔", "〕")).fix(s)

        @staticmethod
        def format_nested_bracket(s: str) -> str:
            return NestedCircumfix(("「", "」"), ("『", "』")).fix(s)

        @staticmethod
        def swap_tabs(s: str) -> str:
            lines = s.splitlines()
            if len(lines) < 1:
                return s
            swapped = []
            for line in lines:
                ss = line.split("\t")
                ss.insert(0, ss.pop())
                swapped.append("\t".join(ss))
            return "\n".join(swapped)

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

    keymap.cutsom_clipboard_formatter = {}

    class ClipboardFormatMenu:
        @staticmethod
        def set_formatter(binding: dict) -> None:
            for menu, func in binding.items():
                keymap.cutsom_clipboard_formatter[menu] = func

        @staticmethod
        def invoke_replacer(search: str, replace_to: str) -> Callable:
            reg = re.compile(search)

            def _replacer(s: str) -> str:
                return reg.sub(replace_to, s)

            return _replacer

        @classmethod
        def set_replacer(cls, binding: dict) -> None:
            for menu, args in binding.items():
                keymap.cutsom_clipboard_formatter[menu] = cls.invoke_replacer(*args)

    ClipboardFormatMenu.set_formatter(
        {
            "swap tabs": FormatTools.swap_tabs,
            "as code": lambda c: f"`{c}`",
            "trim space on line head": FormatTools.trim_space_on_line_head,
            "my markdown frontmatter": lambda _: md_frontmatter(),
            "to lowercase": lambda c: c.lower(),
            "to uppercase": lambda c: c.upper(),
            "to slack feed subscribe": lambda c: "/feed subscribe {}".format(c),
            "to slack feed remove": lambda c: "/feed remove {}".format(c),
            "to list": FormatTools.to_list,
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
            "fix nested paren": FormatTools.format_nested_paren,
            "fix nested bracket": FormatTools.format_nested_bracket,
            "zoom invitation": format_zoom_invitation,
            "remove whitespaces": StrCleaner.remove_whitespace,
        }
    )
    ClipboardFormatMenu.set_replacer(
        {
            "escape backslash": (r"\\", r"\\\\"),
            "escape double-quotation": (r"\"", r'\\"'),
            "remove double-quotation": (r'"', ""),
            "remove single-quotation": (r"'", ""),
            "remove linebreak": (r"\r?\n", ""),
            "to sigle line": (r"\r?\n", ""),
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

    def fzfmenu() -> None:
        if not check_fzf():
            balloon("cannot find fzf on PC.")
            return

        if not ClipHandler.get_string():
            balloon("no text in clipboard.")
            return

        table = keymap.cutsom_clipboard_formatter

        def _fzf(job_item: ckit.JobItem) -> None:
            job_item.func = None

            proc = subprocess.Popen(
                ["fzf.exe", "--no-mouse", "--margin=1"],
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

            job_item.func = table.get(result, None)

        def _finished(job_item: ckit.JobItem) -> None:
            if job_item.func:
                ClipHandler().paste(None, job_item.func)

        subthread_run(_fzf, _finished, True)

    keymap_global["U1-Z"] = suppress_binded_key(fzfmenu)


def configure_ListWindow(window: ListWindow) -> None:
    window.keymap["J"] = window.command_CursorDown
    window.keymap["K"] = window.command_CursorUp
    window.keymap["C-J"] = window.command_CursorPageDown
    window.keymap["C-K"] = window.command_CursorPageUp
    window.keymap["L"] = window.command_Enter
    for mod in ["", "S-", "C-", "C-S-"]:
        for key in ["L", "Space"]:
            window.keymap[mod + key] = window.command_Enter

    def to_top_of_list():
        if window.isearch:
            return
        window.select = 0
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["A"] = to_top_of_list

    def to_end_of_list():
        if window.isearch:
            return
        window.select = len(window.items) - 1
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["E"] = to_end_of_list
