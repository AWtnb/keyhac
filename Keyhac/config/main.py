import fnmatch
import os
import re
import subprocess
import unicodedata
import urllib.parse
import webbrowser
from collections.abc import Callable
from enum import Enum
from winreg import HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, OpenKey, QueryValueEx

import ckit  # type: ignore
import pyauto  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from .tools import clipboard as cb_tool
from .tools import ime as ime_tool
from .tools import sender as sender_tool
from .tools import subthread as subthread_tool
from .tools import virtual_finger as vf_tool
from .tools.clipboard import (
    FIFOStack,
    invoke_clean_paster,
    invoke_quote_paster,
    smart_copy,
    smart_paste,
)
from .tools.common import (
    CallbackFunc,
    balloon,
    check_fzf,
    delay,
    get_now,
    is_browser,
    is_global_target,
    is_keyhac_console,
    open_vscode,
    shell_exec,
    smart_check_path,
)
from .tools.format_str import (
    as_single_line,
    remove_whitespace,
    simple_quote,
    skip_blank_line,
)


def setup(keymap) -> None:

    vf_tool.setup(keymap)
    subthread_tool.keymap = keymap
    ime_tool.keymap = keymap
    sender_tool.keymap = keymap
    cb_tool.setup(keymap)

    ################################
    # general setting
    ################################

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

    # keymap working on any window
    keymap_global = keymap.defineWindowKeymap(check_func=is_global_target)

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
            "LC-U1-T": ("LWin-T"),
            # send n and space
            "LS-U0-N": ("N", "N", "Space"),
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
            "U0-U": ("S-BackSlash"),
            "U1-S": ("Slash"),
            "U0-4": ("S-4", "S-BackSlash"),
            "U0-Enter": ("Period"),
            "U0-Z": ("Minus"),
            "U1-X": ("S-1"),
            # Insert line
            "U0-I": ("End", "Enter"),
            "S-U0-I": ("Home", "Enter", "Up"),
            # Context menu
            "U0-C": ("Apps"),
            "S-U0-C": ("S-Apps"),
            # rename
            "U0-N": ("F2"),
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
        },
    )

    keymap_global["U1-Up"] = keymap.MouseWheelCommand(1.0)
    keymap_global["U1-Down"] = keymap.MouseWheelCommand(-1.0)
    keymap_global["U1-Left"] = keymap.MouseHorizontalWheelCommand(-1.0)
    keymap_global["U1-Right"] = keymap.MouseHorizontalWheelCommand(1.0)

    def safe_close() -> None:
        finger = vf_tool.VirtualFinger(0)
        close_key = finger.compile("A-F4")

        def _wait(_) -> None:
            delay(200)

        def _close(_) -> None:
            finger.send_compiled(*close_key)

        subthread_tool.run(_wait, _close)

    keymap_global["C-Q"] = safe_close

    def bind_ime_control() -> None:
        control = ime_tool.ImeControl(0)
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
            "U1-U": control.start_skk_conv_suffix,
        }.items():
            keymap_global[key] = func

    bind_ime_control()

    keymap.fifo_stack = FIFOStack()
    keymap_global["LC-LS-U0-X"] = keymap.fifo_stack.toggle

    keymap_global["LC-LS-U0-F"] = lambda: keymap.fifo_stack.bulk_register(
        cb_tool.get_string()
    )

    keymap_global["LC-C"] = smart_copy(False)
    keymap_global["LC-X"] = smart_copy(True)
    keymap_global["LC-V"] = smart_paste(False)
    keymap_global["U0-V"] = smart_paste(True)

    ################################
    # custom hotkey
    ################################

    def bind_cleanup_paster(km: WindowKeymap, key: str) -> None:
        for mod1, no_space in {
            "": False,
            "C-": True,
        }.items():
            for mod2, no_break in {
                "": False,
                "S-": True,
            }.items():
                km[mod1 + mod2 + key] = invoke_clean_paster(no_space, no_break)

    keymap_global["U1-V"] = keymap.defineMultiStrokeKeymap()
    bind_cleanup_paster(keymap_global["U1-V"], "V")

    # paste with quote mark
    keymap_global["U1-Q"] = invoke_quote_paster(simple_quote)
    keymap_global["LC-U1-Q"] = invoke_quote_paster(as_single_line)
    keymap_global["LS-U1-Q"] = invoke_quote_paster(skip_blank_line)

    # open url in browser
    def open_selected_url() -> None:
        def _open(job_item: ckit.JobItem) -> None:
            if job_item.copied:
                u = job_item.copied
            else:
                u = job_item.origin
            u = u.strip()
            if u.startswith("http"):
                webbrowser.open(u)
            else:
                balloon(keymap, f"invalid path: {u}")

        cb_tool.Manager().after_register(_open)

    keymap_global["C-U0-O"] = open_selected_url

    ################################
    # config keys
    ################################

    def reload_config() -> None:
        def _wait(_) -> None:
            delay(60)

        def _reload(_) -> None:
            ckit.JobQueue.cancelAll()
            keymap.configure()
            keymap.updateKeymap()
            ts = get_now().strftime("%Y-%m-%d %H:%M:%S")
            balloon(keymap, f"{ts} reloaded config.py")

        subthread_tool.run(_wait, _reload)

    keymap_global["U1-F12"] = reload_config

    def open_keyhac_repo() -> None:
        config_path = os.path.expandvars(r"${APPDATA}\Keyhac")
        if not smart_check_path(config_path):
            balloon(keymap, f"config not found: {config_path}")
            return

        dir_path = config_path
        if (real_path := os.path.realpath(config_path)) != dir_path:
            dir_path = os.path.dirname(real_path)

        def _open(_) -> None:
            result = open_vscode(dir_path)
            if not result:
                shell_exec(dir_path)

        subthread_tool.run(_open)

    keymap.editor = lambda _: open_keyhac_repo()

    keymap_global["U0-F12"] = open_keyhac_repo

    # clipboard menu
    keymap_global["LC-LS-X"] = keymap.command_ClipboardList

    ################################
    # set window position
    ################################

    def bind_window_mover(km: WindowKeymap) -> None:
        for key, delta in {
            "Left": (-10, 0),
            "Right": (+10, 0),
            "Up": (0, -10),
            "Down": (0, +10),
        }.items():
            x, y = delta
            for mod, scale in {"": 15, "S-": 5, "C-": 5, "S-C-": 1}.items():
                km[mod + "U0-" + key] = keymap.MoveWindowCommand(x * scale, y * scale)

    bind_window_mover(keymap_global)

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

        def move_edge(self, toward: RectEdge, delta: int) -> list[int]:
            r = [
                self.left,
                self.top,
                self.right,
                self.bottom,
            ]
            opposite = (toward.value + 2) % 4
            r[opposite] = r[toward.value] + delta
            return r

        def resize(self, scale: float, toward: RectEdge) -> list[int]:
            if toward in [RectEdge.left, RectEdge.right]:
                dim = self.width
            else:
                dim = self.height
            delta = int(dim * scale)
            if toward in [RectEdge.right, RectEdge.bottom]:
                delta = delta * -1
            return self.move_edge(toward, delta)

    def bind_window_snapper(km: WindowKeymap) -> None:
        altkey_stat = ["", "LA-", "RA-"]
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

        for idx, alt in enumerate(altkey_stat):
            for area_mod, scale in scale_mapping.items():
                for key, edge in edge_mapping.items():

                    def _invoke(
                        i: int = idx, s: float = scale, e: RectEdge = edge
                    ) -> CallbackFunc:
                        def __get_new_rect() -> list[int]:
                            infos = pyauto.Window.getMonitorInfo()
                            infos.sort(key=lambda info: info[2] != 1)
                            target = infos[i]
                            monitor_work_rect = Rect(*target[1])
                            return monitor_work_rect.resize(s, e)

                        def __snap() -> None:
                            wnd = keymap.getTopLevelWindow()
                            if not wnd or is_keyhac_console(wnd):
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

                            subthread_tool.run(_job_snap, _job_finished)

                        return __snap

                    km[alt + area_mod + key] = _invoke()

    bind_window_snapper(keymap_global["U1-M"])

    def bind_maximized_window_snapper() -> None:
        for key in ["0", "1", "2"]:
            monitor_idx = int(key)

            def _snap(mi: int = monitor_idx) -> None:
                infos = pyauto.Window.getMonitorInfo()
                infos.sort(key=lambda info: info[2] != 1)
                try:
                    target = infos[mi][1]
                except IndexError:
                    return

                def __snap(job_item: ckit.JobItem) -> None:
                    job_item.wnd = None

                    wnd = keymap.getTopLevelWindow()
                    if not wnd or is_keyhac_console(wnd):
                        return
                    if wnd.isMaximized():
                        wnd.restore()
                        delay()
                    wnd.setRect(target)
                    job_item.wnd = wnd

                def __maximize(job_item: ckit.JobItem) -> None:
                    job_item.wnd.maximize()

                subthread_tool.run(__snap, __maximize)

            keymap_global["U1-M"][str(key)] = _snap

    bind_maximized_window_snapper()

    class WndShrinker:
        @staticmethod
        def invoke_snapper(toward: RectEdge) -> CallbackFunc:
            def _snapper() -> None:
                def __snap(_) -> None:
                    wnd = keymap.getTopLevelWindow()
                    rect = wnd.getRect()

                    resized = Rect(*rect).resize(0.5, toward)
                    if wnd.isMaximized():
                        wnd.restore()
                        delay()
                    wnd.setRect(resized)

                subthread_tool.run(__snap)

            return _snapper

        @classmethod
        def bind(cls, km: WindowKeymap) -> None:
            for key, toward in {
                "H": RectEdge.left,
                "L": RectEdge.right,
                "K": RectEdge.top,
                "J": RectEdge.bottom,
            }.items():
                km["U1-" + key] = cls.invoke_snapper(toward)

    WndShrinker.bind(keymap_global["U1-M"])

    class MonitorCenterAvoider:
        @staticmethod
        def get_border_x(wnd: pyauto.Window) -> int:
            x = -1
            rect = Rect(*wnd.getRect())
            center_x = int((rect.left + rect.right) / 2)
            monitors = pyauto.Window.getMonitorInfo()
            for monitor in monitors:
                monitor_rect = Rect(*monitor[1])
                if monitor_rect.left <= center_x <= monitor_rect.right:
                    x = int((monitor_rect.left + monitor_rect.right) / 2)
                    break
            return x

        @classmethod
        def invoke_avoider(cls, show_left: bool) -> CallbackFunc:
            def _avoider() -> None:
                def _snap(_) -> None:
                    wnd = keymap.getTopLevelWindow()
                    if is_keyhac_console(wnd) or wnd.isMaximized():
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

                subthread_tool.run(_snap)

            return _avoider

    keymap_global["U1-M"]["OpenBracket"] = MonitorCenterAvoider.invoke_avoider(False)
    keymap_global["U1-M"]["CloseBracket"] = MonitorCenterAvoider.invoke_avoider(True)

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

    keymap_global["O-RCtrl"] = CursorPos.snap
    keymap_global["O-RShift"] = CursorPos.snap_to_center

    ################################
    # input customize
    ################################

    keymap_global["Yen"] = sender_tool.SKKSender().invoke_emitThen(
        ime_tool.ImeStatus.off, "Yen"
    )

    keymap_global["U0-P"] = sender_tool.SKKSender().under_kanamode("・")

    keymap_global["U0-AtMark"] = sender_tool.SKKSender().invoke_emitThen(
        ime_tool.ImeStatus.off, "LS-AtMark", "LS-AtMark", "Left"
    )
    keymap_global["U0-5"] = sender_tool.SKKSender().invoke_emitThen(
        ime_tool.ImeStatus.off, "S-7", "S-5", "S-5", "S-7", "Left", "Left"
    )

    # select-to-left with ime control
    keymap_global["U1-B"] = sender_tool.SKKSender().under_kanamode("S-Left")
    keymap_global["LS-U1-B"] = sender_tool.SKKSender().under_kanamode("S-Right")
    keymap_global["U1-Space"] = sender_tool.SKKSender().under_kanamode("C-S-Left")
    keymap_global["U1-4"] = sender_tool.SKKSender().under_kanamode(
        ime_tool.SKKKey.convpoint, "S-4", "Tab"
    )

    def bind_fullwidth_sender() -> None:
        sender = sender_tool.SKKSender()
        for key, symbol in {
            "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
            "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
            "S-U0-Period": "\uff0e",  # FULLWIDTH PERIOD
        }.items():
            keymap_global[key] = sender.invoke(
                sender.control.to_skk_full_latin, symbol, ime_tool.SKKKey.kana
            )

    bind_fullwidth_sender()

    def bind_fullwidth_circumfix_sender() -> None:
        sender = sender_tool.SKKSender()
        for key, pair in {
            "U0-8": ["\u300e", "\u300f"],  # WHITE CORNER BRACKET 『』
            "U0-9": ["\u3010", "\u3011"],  # BLACK LENTICULAR BRACKET 【】
            "U0-OpenBracket": ["\u300c", "\u300d"],  # CORNER BRACKET 「」
            "U1-2": ["\u201c", "\u201d"],  # DOUBLE QUOTATION MARK “”
            "U1-7": ["\u2018", "\u2019"],  # SINGLE QUOTATION MARK ‘’
            "U1-8": ["\uff08", "\uff09"],  # FULLWIDTH PARENTHESIS （）
            "U0-Y": ["\u3008", "\u3009"],  # ANGLE BRACKET 〈〉
            "U1-Y": ["\u300a", "\u300b"],  # DOUBLE ANGLE BRACKET 《》
            "U1-T": ["\u3014", "\u3015"],  # TORTOISE BRACKET 〔〕
            "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
        }.items():
            keymap_global[key] = sender.invoke(
                sender.control.to_skk_full_latin, *pair, "Left", ime_tool.SKKKey.kana
            )

    bind_fullwidth_circumfix_sender()

    keymap_global["U0-M"] = keymap.defineMultiStrokeKeymap()

    def replace_last_nchar(km: WindowKeymap, newstr: str) -> None:
        for n in "0123":
            seq = ["Back"] * int(n) + [newstr, ime_tool.SKKKey.toggle_vk]
            km[n] = sender_tool.DirectSender().invoke(*seq)

    replace_last_nchar(keymap_global["U0-M"], "先生")

    def bind_direct_sender(
        km: WindowKeymap, binding: dict[str, tuple[str, ...]]
    ) -> None:
        sender = sender_tool.DirectSender()
        for key, sent in binding.items():
            km[key] = sender.invoke(*sent)

    bind_direct_sender(
        keymap_global,
        {
            "Decimal": ("Period",),
            "U0-1": ("S-1",),
            "U0-Colon": ("Colon",),
            "U0-Slash": ("Slash",),
            "U1-Minus": ("Minus",),
            "LC-U0-U": ("Minus",),
            "U0-Comma": ("Comma",),
            "U0-Period": ("Period",),
            "S-U0-Enter": ("U-Shift", "Period"),
            "U0-Tab": ("Period", "BackSlash"),
            "U1-Tab": ("Period", "Period", "BackSlash"),
            "S-U0-8": ("U-Shift", "Minus", "Space", ime_tool.SKKKey.toggle_vk),
            "U1-1": ("1.", "Space", ime_tool.SKKKey.toggle_vk),
            "S-U0-SemiColon": ("U-Shift", "SemiColon"),
            "U0-T": ("</>", "Left", "S-Left"),
        },
    )

    def bind_direct_sender_circumfix(
        km: WindowKeymap, binding: dict[str, list[str]]
    ) -> None:
        sender = sender_tool.DirectSender()
        for key, circumfix in binding.items():
            _, suffix = circumfix
            sequence = circumfix + ["Left"] * len(suffix)
            km[key] = sender.invoke(*sequence)

    bind_direct_sender_circumfix(
        keymap_global,
        {
            "U0-CloseBracket": ["[", "]"],
            "U1-9": ["(", ")"],
            "S-U0-9": ['("', '")'],
            "U1-CloseBracket": ["{", "}"],
        },
    )

    ################################
    # web search
    ################################

    class Radicals:
        mapping = {
            "\u2e92": "\u5df3",
            "\u2e9f": "\u6bcd",
            "\u2ea0": "\u6c11",
            "\u2ec4": "\u897f",
            "\u2ed1": "\u9577",
            "\u2ed8": "\u9752",
            "\u2ee4": "\u9b3c",
            "\u2ee8": "\u9ea6",
            "\u2ee9": "\u9ec4",
            "\u2eeb": "\u6589",
            "\u2eed": "\u6b6f",
            "\u2eef": "\u7adc",
            "\u2ef2": "\u4e80",
        }

        @classmethod
        def fix(cls, s: str) -> str:
            return s.translate(str.maketrans(cls.mapping))

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
        def __init__(self) -> None:
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
            self._query = Radicals.fix(KangxiRadicals.fix(self._query))

        def remove_honorific(self) -> None:
            for honor in ["先生", "様"]:
                self._query = self._query.replace(honor, " ")

        def remove_editorial_style(self) -> None:
            for honor in [
                "監修",
                "共著",
                "共編著",
                "編著",
                "共編",
                "分担執筆",
                "et al.",
            ]:
                self._query = self._query.replace(honor, " ")

        def remove_hiragana(self) -> None:
            self._query = re.sub(r"[\u3041-\u3093]", " ", self._query)

        def encode(self, strict: bool = False) -> str:
            words = []
            for word in self.noise_mapping.cleanup(self._query).split(" "):
                if len(word):
                    if strict:
                        words.append(f'"{word}"')
                    else:
                        words.append(word)
            return urllib.parse.quote(" ".join(words))

    class WebSearcher:
        def __init__(self, uri_mapping: dict) -> None:
            self._uri_mapping = uri_mapping

        @staticmethod
        def invoke(
            uri: str, strict: bool = False, strip_hiragana: bool = False
        ) -> CallbackFunc:
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

                cb_tool.Manager().after_register(_search)

            return _searcher

        def bind(self, km: WindowKeymap) -> None:
            for shift_key in ("", "S-"):
                for ctrl_key in ("", "C-"):
                    is_strict = 0 < len(shift_key)
                    strip_hiragana = 0 < len(ctrl_key)
                    trigger_key = shift_key + ctrl_key + "U0-S"
                    km[trigger_key] = keymap.defineMultiStrokeKeymap()
                    for key, uri in self._uri_mapping.items():
                        km[trigger_key][key] = self.invoke(
                            uri, is_strict, strip_hiragana
                        )

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
    ).bind(keymap_global)

    ################################
    # activate window
    ################################

    class SystemBrowser:
        prog_id: str = ""
        commandline: str = ""

        def __init__(self) -> None:
            self.set_prog_id()
            self.set_commandline()

        def set_prog_id(self) -> None:
            registry_paths = [
                r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoiceLatest\ProgId",
                r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice",
            ]

            for path in registry_paths:
                try:
                    with OpenKey(HKEY_CURRENT_USER, path) as key:
                        self.prog_id = str(QueryValueEx(key, "ProgId")[0])
                        return
                except Exception as e:
                    print(e)
                    print(f"Failed to get ProgId by registry `{path}`")
                    if path != registry_paths[-1]:
                        print("Try next path...")

        def set_commandline(self) -> None:
            if self.prog_id == "":
                return
            registry_path = os.path.join(self.prog_id, "shell", "open", "command")
            try:
                with OpenKey(HKEY_CLASSES_ROOT, registry_path) as key:
                    self.commandline = str(QueryValueEx(key, "")[0])
            except Exception as e:
                print(e)
                print(f"Failed to get ProgId by registry `{registry_path}`")

        def get_exe_path(self) -> str:
            if self.commandline == "" or self.prog_id == "":
                return ""
            c = self.commandline
            e = ".exe"
            return c[: c.find(e) + len(e)].strip('"')

        def get_exe_name(self) -> str:
            if self.commandline == "" or self.prog_id == "":
                return ""
            _, name = os.path.split(self.get_exe_path())
            return name

        def get_wnd_class(self) -> str:
            return {
                "chrome.exe": "Chrome_WidgetWin_1",
                "vivaldi.exe": "Chrome_WidgetWin_1",
                "firefox.exe": "MozillaWindowClass",
            }.get(self.get_exe_name(), "Chrome_WidgetWin_1")

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
            if self.class_name and not fnmatch.fnmatch(
                wnd.getClassName(), self.class_name
            ):
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
        def invoke(
            exe_name: str, class_name: str = "", exe_path: str | CallbackFunc = ""
        ) -> CallbackFunc:
            def _executor() -> None:
                def _activate(job_item: ckit.JobItem) -> None:
                    job_item.result = None
                    delay(40)
                    scanner = WndScanner(exe_name, class_name)
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
                    if job_item.result is not None:
                        if not job_item.result:
                            vf_tool.VirtualFinger().send("LCtrl-LAlt-Tab")

                subthread_tool.run(_activate, _finished, True)

            return _executor

        @classmethod
        def bind(cls, wnd_keymap: WindowKeymap, remap_table: dict = {}) -> None:
            for key, params in remap_table.items():
                func = cls.invoke(*params)
                wnd_keymap[key] = func

    PseudoCuteExec.bind(
        keymap_global,
        {
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
        },
    )

    keymap_global["U1-C"] = keymap.defineMultiStrokeKeymap()
    PseudoCuteExec.bind(
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
            "V": ("Code.exe", "Chrome_WidgetWin_1", open_vscode),
            "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1"),
            "M": (
                "Mery.exe",
                "TChildForm",
                r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
            ),
            "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
        },
    )

    def fuzzy_window_switcher() -> None:
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
                    vf_tool.VirtualFinger().send("LCtrl-LAlt-Tab")

        subthread_tool.run(_fzf_wnd, _finished, True)

    keymap_global["U1-E"] = fuzzy_window_switcher

    def invoke_draft() -> None:
        def _invoke(_) -> None:
            shell_exec(r"${USERPROFILE}\Personal\draft.txt")

        subthread_tool.run(_invoke)

    keymap_global["LS-LC-U1-M"] = invoke_draft

    def search_on_browser() -> None:
        finger = vf_tool.VirtualFinger(20)
        if keymap.getWindow().getProcessName() == keymap.default_browser.get_exe_name():
            finger.send("C-T")
            return

        def _activate(job_item: ckit.JobItem) -> None:
            delay()
            job_item.result = None
            scanner = WndScanner(
                keymap.default_browser.get_exe_name(),
                keymap.default_browser.get_wnd_class(),
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

    keymap_global["U0-Q"] = search_on_browser

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"
    keymap_browser["LC-F"] = sender_tool.SKKSender(40).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-F"
    )
    keymap_browser["LC-K"] = sender_tool.SKKSender(40).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-K"
    )

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None

    # rsturio
    keymap_rstudio = keymap.defineWindowKeymap(exe_name="rstudio.exe")
    keymap_rstudio["U0-Minus"] = sender_tool.DirectSender().invoke("S-Comma", "Minus")

    # slack
    keymap_slack = keymap.defineWindowKeymap(
        exe_name="slack.exe", class_name="Chrome_WidgetWin_1"
    )
    keymap_slack["C-K"] = sender_tool.SKKSender().invoke_emitThen(
        ime_tool.ImeStatus.off, "C-K"
    )
    keymap_slack["F3"] = keymap_slack["C-K"]
    keymap_slack["C-E"] = keymap_slack["C-K"]
    keymap_slack["F1"] = sender_tool.DirectSender().invoke("S-SemiColon", "Colon")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")
    keymap_vscode["U0-Slash"] = "C-Slash", "A-S-Down", "C-Slash"

    def remap_vscode(*keys: str) -> None:
        sender = sender_tool.SKKSender()
        for key in keys:
            keymap_vscode[key] = sender.invoke_emitThen(ime_tool.ImeStatus.off, key)

    remap_vscode(
        "C-E",
        "C-F",
        "C-T",
        "C-S-F",
        "C-S-E",
        "C-S-O",
        "C-S-G",
        "RC-RS-X",
        "C-0",
        "C-S-P",
        "C-A-B",
        "C-A-AtMark",
        "C-1",
        "C-2",
        "C-S-Enter",
        "S-Enter",
    )

    # mery
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")

    def remap_mery(binding: dict) -> None:
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

    # Kiri
    keymap_kiri = keymap.defineWindowKeymap(
        exe_name="KIRI10.exe", class_name="TblCommCtrl"
    )
    keymap_kiri["F2"] = "F2", "End"
    keymap_kiri["U0-N"] = keymap_kiri["F2"]

    keymap_kiri_edit = keymap.defineWindowKeymap(
        exe_name="KIRI10.exe", class_name="Edit"
    )
    keymap_kiri_edit["C-Enter"] = "F4", "Down"
    keymap_kiri_edit["LC-U0-Space"] = keymap_kiri_edit["C-Enter"]

    # smooth csv
    keymap_smoothcsv = keymap.defineWindowKeymap(
        exe_name="msedgewebview2.exe",
        class_name="Chrome_WidgetWin_1",
        window_text="tauri.localhost",
    )
    keymap_smoothcsv["C-S-F"] = sender_tool.SKKSender(80).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-S-F", "C-A"
    )
    keymap_smoothcsv["S-Space"] = sender_tool.DirectSender().invoke("S-Space")
    keymap_smoothcsv["S-U0-N"] = lambda: vf_tool.VirtualFinger(20).send("F2", "Home")

    def copy_and_unselect_line() -> None:
        finger = vf_tool.VirtualFinger()
        taps = vf_tool.VirtualFinger().compile("Up", "Down")

        def _unselect(_) -> None:
            finger.send_compiled(*taps)

        cb_tool.Manager().after_register(_unselect)

    keymap_smoothcsv["U1-C"] = copy_and_unselect_line

    # sumatra PDF
    keymap_sumatra = keymap.defineWindowKeymap(
        check_func=lambda wnd: wnd.getProcessName() == "SumatraPDF.exe"
    )
    keymap_sumatra["O-LCtrl"] = "Esc", "Esc", "C-Home", "C-F"

    # sumatra PDF (not focused on inputbox)
    keymap_sumatra_view = keymap.defineWindowKeymap(
        check_func=(
            lambda wnd: (
                wnd.getProcessName() == "SumatraPDF.exe"
                and wnd.getClassName() != "Edit"
            )
        )
    )

    def sumatra_view_key() -> None:
        sender = sender_tool.DirectSender()
        for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            keymap_sumatra_view[key] = sender.invoke(key)

    sumatra_view_key()

    keymap_sumatra_view["H"] = "C-S-Tab"
    keymap_sumatra_view["L"] = "C-Tab"

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")  # noqa: F841

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    keymap_ppt["O-(236)"] = ime_tool.ImeControl(40).to_skk_abbrev

    # excel
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")

    def select_all() -> None:
        finger = vf_tool.VirtualFinger()
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

    keymap_global["U1-W"] = lambda: cb_tool.Manager().paste(
        format_func=CharWidth(True).to_full_letter
    )
    keymap_global["LS-U1-W"] = lambda: cb_tool.Manager().paste(
        format_func=CharWidth(True).to_half_letter
    )

    def format_zoom_invitation(s: str) -> str:
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
        def skip_blank_line(s: str) -> str:
            lines = s.strip().splitlines()
            return "\n".join([line for line in lines if line.strip()])

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
            return "\n".join(["- " + line for line in lines])

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
                    ss.append(f"{m.group(1)}-{m.group(2)}\t{m.group(3)}")
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
                    return f"\u201c{mo.group(1)}\u201d"
                return f"\u2018{mo.group(1)}\u2019"

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

            def _split(s: str) -> list[str]:
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

    def invoke_comment_remover(symbol: str) -> Callable[[str], str]:
        def _remover(s: str) -> str:
            return "\n".join(
                [line for line in s.splitlines() if not line.strip().startswith(symbol)]
            )

        return _remover

    keymap.cutsom_clipboard_formatter = {}

    class ClipboardFormatMenu:
        @staticmethod
        def set_formatter(binding: dict) -> None:
            for menu, func in binding.items():
                keymap.cutsom_clipboard_formatter[menu] = func

        @staticmethod
        def invoke_replacer(search: str, replace_to: str) -> Callable[[str], str]:
            reg = re.compile(search)

            def _replacer(s: str) -> str:
                return reg.sub(replace_to, s)

            return _replacer

        @classmethod
        def set_replacer(cls, binding: dict) -> None:
            for menu, args in binding.items():
                keymap.cutsom_clipboard_formatter[menu] = cls.invoke_replacer(*args)

        @staticmethod
        def invoke_line_jointer(sep: str) -> Callable[[str], str]:
            def _jointer(s: str) -> str:
                return sep.join(s.splitlines())

            return _jointer

        @classmethod
        def set_line_jointer(cls, binding: dict) -> None:
            for name, sep in binding.items():
                menu = f"Join lines with {name}"
                keymap.cutsom_clipboard_formatter[menu] = cls.invoke_line_jointer(sep)

    ClipboardFormatMenu.set_formatter(
        {
            "to codeblock": lambda c: f"```\n{c}\n```\n",
            "swap tabs": FormatTools.swap_tabs,
            "trim space on line head": FormatTools.trim_space_on_line_head,
            "my markdown frontmatter": lambda _: md_frontmatter(),
            "FIFO: join items with Tab": lambda _: keymap.fifo_stack.join_items("\t"),
            "FIFO: join items with LineBreak": lambda _: keymap.fifo_stack.join_items(
                "\n"
            ),
            "FIFO: bulk paste": lambda _: keymap.fifo_stack.bulk_paste("\n"),
            "FIFO: bulk paste (join with space)": lambda _: (
                keymap.fifo_stack.bulk_paste(" ")
            ),
            "to lowercase": lambda c: c.lower(),
            "to uppercase": lambda c: c.upper(),
            "to slack feed subscribe": lambda c: f"/feed subscribe {c}",
            "to slack feed remove": lambda c: f"/feed remove {c}",
            "to list": FormatTools.to_list,
            "to deepl-friendly": FormatTools.to_deepl_friendly,
            "swap abbreviation around colon": FormatTools.swap_abbreviation,
            "colon to double-dash": FormatTools.colon_to_doubledash,
            "insert blank line": FormatTools.insert_blank_line,
            "remove blank line": FormatTools.skip_blank_line,
            "fix dumb quotation": FormatTools.fix_dumb_quotation,
            "fix KANGXI RADICALS": KangxiRadicals.fix,
            "fix paren inside bracket": FormatTools.fix_paren_inside_bracket,
            "to double bracket": FormatTools.to_double_bracket,
            "to single bracket": FormatTools.to_single_bracket,
            "TSV to markdown table": FormatTools.mdtable_from_tsv,
            "split postalcode and address": FormatTools.split_postalcode,
            "decode url": FormatTools.decode_url,
            "encode url": FormatTools.encode_url,
            "to halfwidth": CharWidth().to_half_letter,
            "to halfwidth (including symbols)": CharWidth(True).to_half_letter,
            "to halfwidth symbols": CharWidth.to_half_symbol,
            "to halfwidth bracktets": CharWidth.to_half_brackets,
            "to fullwidth": CharWidth().to_full_letter,
            "to fullwidth (including symbols)": CharWidth(True).to_full_letter,
            "to fullwidth symbols": CharWidth.to_full_symbol,
            "to fullwidth bracktets": CharWidth.to_full_brackets,
            "trim honorific": FormatTools.trim_honorific,
            "fix nested paren": FormatTools.format_nested_paren,
            "fix nested bracket": FormatTools.format_nested_bracket,
            "zoom invitation": format_zoom_invitation,
            "remove whitespaces": remove_whitespace,
            "remove javascript comment line": invoke_comment_remover("// "),
            "remove python comment line": invoke_comment_remover("# "),
        }
    )
    ClipboardFormatMenu.set_replacer(
        {
            "backslash to slash": (r"\\", "/"),
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
            "fix msword-bullet": (
                r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009",
                "\u30fb",
            ),
            "remove msword-bullet": (
                r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009",
                "",
            ),
            "to curly-comma (\uff0c)": (r"\u3001", "\uff0c"),
            "to japanese-comma (\u3001)": (r"\uff0c", "\u3001"),
            "shorten amazon url": (
                r"^.+amazon\.co\.jp/.+dp/(.{10}).*",
                r"https://www.amazon.jp/dp/\1",
            ),
        }
    )

    ClipboardFormatMenu.set_line_jointer(
        {
            "Half-Comma": ",",
            "Dot": "・",
            "Tab": "\t",
            "Slash": "／",
            "Pipe": "|",
        }
    )

    def fzfmenu() -> None:
        if not check_fzf():
            balloon(keymap, "cannot find fzf on PC.")
            return

        if not cb_tool.get_string():
            balloon(keymap, "no text in clipboard.")
            return

        table = keymap.cutsom_clipboard_formatter

        def _fzf(job_item: ckit.JobItem) -> None:
            job_item.func = None
            delay()

            proc = subprocess.Popen(
                ["fzf.exe", "--no-mouse", "--margin=1"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
            )
            try:
                if proc.stdin:
                    for k in table.keys():
                        proc.stdin.write(k + "\n")
                    proc.stdin.close()
            except Exception as e:
                balloon(keymap, e)
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
                cb_tool.Manager().paste(None, job_item.func)

        subthread_tool.run(_fzf, _finished, True)

    keymap_global["U1-Z"] = fzfmenu
