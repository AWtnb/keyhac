import datetime
import os
import fnmatch
import re
import time
import urllib.parse
from pathlib import Path
from winreg import HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, OpenKey, QueryValueEx

import pyauto
from keyhac import *


def configure(keymap):

    ################################
    # general setting
    ################################

    def resolve_path(rel:str, root:str=None) -> str:
        if not root:
            root = os.environ.get("USERPROFILE")
        return str(Path(root, rel))

    def test_path(s:str) -> bool:
        return Path(s).exists()

    def get_filer_path() -> str:
        if test_path(tablacus_path := resolve_path(r"Dropbox\portable_apps\tablacus\TE64.exe")):
            return tablacus_path
        return "explorer.exe"

    def get_editor_path() -> str:
        if test_path(vscode_path := resolve_path(r"scoop\apps\vscode\current\Code.exe")):
            return vscode_path
        return "notepad.exe"
    keymap.editor = get_editor_path()

    # console theme
    keymap.setFont("HackGen", 16)
    keymap.setTheme("black")

    # user modifier
    keymap.replaceKey("(29)", 235) # "muhenkan" => 235
    keymap.replaceKey("(28)", 236) # "henkan" => 236
    keymap.defineModifier(235, "User0") # "muhenkan" => "U0"
    keymap.defineModifier(236, "User1") # "henkan" => "U1"

    # enable clipbard history
    keymap.clipboard_history.enableHook(True)

    # history max size
    keymap.clipboard_history.maxnum = 200
    keymap.clipboard_history.quota = 10*1024*1024

    # quote mark when paste with Ctrl.
    keymap.quote_mark = "> "

    ################################
    # key remap
    ################################

    def is_browser_wnd(wnd:pyauto.Window) -> bool:
        return wnd.getProcessName() in ("chrome.exe", "vivaldi.exe", "firefox.exe")

    # keymap working on any window
    def is_global_target(wnd:pyauto.Window) -> bool:
        return not ( is_browser_wnd(wnd) and wnd.getText().startswith("ESET - ") )
    keymap_global = keymap.defineWindowKeymap(check_func=is_global_target)

    # keyboard macro
    keymap_global["U0-Z"] = keymap.command_RecordPlay
    keymap_global["U0-0"] = keymap.command_RecordToggle
    keymap_global["U1-0"] = keymap.command_RecordClear

    # combination with modifier key
    for mod_mntr in ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-"):
        for key, value in {
            # move cursor
            "H": "Left",
            "J": "Down",
            "K": "Up",
            "L": "Right",
            # Home / End
            "A": "Home",
            "E": "End",
            # Enter
            "Space": "Enter",
        }.items():
            keymap_global[mod_mntr+"U0-"+key] = mod_mntr+value

        for stat in ("D-", "U-"):
            # ignore capslock
            keymap_global[mod_mntr+stat+"Capslock"] = lambda : None
            # ignore "katakana-hiragana-romaji key"
            for vk in list(range(124, 136)) + list(range(240, 243)) + list(range(245, 254)):
                keymap_global[mod_mntr+stat+str(vk)] = lambda : None

    for key, value in {
    # BackSpace / Delete
        "U0-D": ("Delete"),
        "U0-B": ("Back"),
        "C-U0-D": ("C-Delete"),
        "C-U0-B": ("C-Back"),
        "S-U0-D": ("S-End", "Delete"),
        "S-U0-B": ("S-Home", "Delete"),

        # escape
        "O-(235)": ("Esc"),
        "U0-X": ("Esc"),

        # select first suggestion
        "U0-Tab": ("Down", "Enter"),

        # line selection
        "U1-A": ("End", "S-Home"),


        # punctuation
        "U0-Enter": ("Period"),
        "LS-U0-Enter": ("Comma"),
        "LC-U0-Enter": ("Slash"),
        "U1-S": ("Slash"),
        "U1-B": ("Minus"),
        "U0-O": ("Minus"),

        # Re-convert
        "U0-R": ("LWin-Slash"),

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

        "Insert": (lambda : None),

    }.items():
        keymap_global[key] = value

    # for key in "123456789":
    #     keymap_global["D-LWin-"+key] = "LWin-"+key

    ################################
    # functions for custom hotkey
    ################################

    def delay(sec:float=0.05) -> None:
        time.sleep(sec)

    def get_current_clipboard() -> str:
        return getClipboardText() or ""

    def to_non_space(s:str) -> str:
        return s.strip().translate(str.maketrans("", "", "\u200b\u3000\u0009\u0020"))

    def send_keys(*keys:str) -> None:
        keymap.beginInput()
        for key in keys:
            keymap.setInput_FromString(str(key))
        keymap.endInput()

    def send_string(s:str) -> None:
        keymap.beginInput()
        keymap.setInput_Modifier(0)
        for c in str(s):
            keymap.input_seq.append(pyauto.Char(c))
        keymap.endInput()

    def send_input(sequence:tuple, sleep:float=0.01) -> None:
        for elem in sequence:
            delay(sleep)
            try:
                send_keys(elem)
            except:
                send_string(elem)

    def set_ime(mode:int) -> None:
        if keymap.getWindow().getImeStatus() != mode:
            send_keys("(243)")
            delay(0.01)

    def copy_string() -> str:
        setClipboardText("")
        send_keys("C-Insert")
        timeout = 0.2
        retry_time = 0.01
        cb = ""
        while timeout > 0.0:
            if s := getClipboardText():
                cb = s
                break
            delay(retry_time)
            timeout -= retry_time
        return cb

    def paste_string(s:str) -> None:
        setClipboardText(s)
        send_keys("S-Insert")

    def execute_path(s:str, arg:str=None) -> None:
        if s:
            if s.startswith("http") or test_path(s):
                keymap.ShellExecuteCommand(None, s, arg, None)()

    def recover_clipboard(func:callable) -> callable:
        def _executer() -> None:
            backup = get_current_clipboard()
            func()
            if backup:
                setClipboardText(backup)
        return _executer

    def lazy_call(func:callable, delay_msec:int=20) -> callable:
        def _hook_call() -> None:
            keymap.hookCall(func)
        def _executer() -> None:
            # use delayedCall in ckit => https://github.com/crftwr/ckit/blob/2ea84f986f9b46a3e7f7b20a457d979ec9be2d33/ckitcore/ckitcore.cpp#L1998
            keymap.delayedCall(_hook_call, delay_msec)
            keymap._fixFunnyModifierState()
        return _executer


    class KeyPuncher:
        def __init__(self, recover_ime:bool=False, sleep_sec:float=0.01, delay_msec:int=20) -> None:
            self._recover_ime = recover_ime
            self._sleep = sleep_sec
            self._delay_msec = delay_msec
        def invoke(self, *sequence) -> callable:
            def _input() -> None:
                set_ime(0)
                send_input(sequence, self._sleep)
                if self._recover_ime:
                    set_ime(1)
            return lazy_call(_input, self._delay_msec)

    ################################
    # release CapsLock on reload
    ################################

    if pyauto.Input.getKeyState(VK_CAPITAL):
        send_keys("LS-CapsLock")
        print("released CapsLock.")

    ################################
    # custom hotkey
    ################################

    # ime dict tool
    keymap_global["C-F7"] = lazy_call(lambda : execute_path(r"C:\Program Files (x86)\Google\Google Japanese Input\GoogleIMEJaTool.exe", "--mode=word_register_dialog"))

    # listup Window
    keymap_global["U0-W"] = lazy_call(lambda : send_keys("LCtrl-LAlt-Tab"))

    # ime: Japanese / Foreign
    keymap_global["U1-J"] = lambda : set_ime(1)
    keymap_global["U0-F"] = lambda : set_ime(0)
    keymap_global["S-U0-F"] = lambda : set_ime(1)
    keymap_global["S-U1-J"] = lambda : set_ime(0)

    # paste as plaintext
    keymap_global["U0-V"] = lazy_call(lambda : paste_string(get_current_clipboard().strip()))
    keymap_global["U1-V"] = lazy_call(lambda : paste_string(to_non_space(get_current_clipboard())))

    # select last word with ime
    def select_last_word() -> None:
        send_keys("C-S-Left")
        set_ime(1)
    keymap_global["U1-Space"] = select_last_word

    # Non-convert
    def as_alphabet(recover_ime:bool=False) -> callable:
        keys = ["F10"]
        if recover_ime:
            keys.append("Enter")
        else:
            keys.append("(243)")
        def _sender():
            if keymap.getWindow().getImeStatus() == 1:
                send_keys(*keys)
        return _sender

    keymap_global["U1-N"] = as_alphabet(False)
    keymap_global["S-U1-N"] = as_alphabet(True)

    # fix typo
    def fix_previous_typo() -> callable:
        typo_map = {
            "/": ".",
            ".": ",",
            ",": ".",
            ":": ";",
            ";": ":",
            "l": ";",
            "\uff1a": "\uff1b",
            "\uff1b": "\uff1a",
            "\u3001": "\u3002",
            "\u3002": "\u3001",
            "\uff01": "\uff1f",
            "\uff1f": "\uff01",
        }
        def _fixer() -> None:
            send_keys("S-Left")
            cb = copy_string()
            if cb:
                sent = typo_map.get(cb, "")
                if sent:
                    send_string(typo_map[cb])
        return lazy_call(recover_clipboard(_fixer))
    keymap_global["U0-Back"] = fix_previous_typo()
    keymap_global["U0-BackSlash"] = fix_previous_typo()

    # count chars
    def count_chars() -> None:
        cb = copy_string()
        if cb:
            total = len(cb)
            lines = len(cb.strip().splitlines())
            net = len(to_non_space(cb.strip()))
            t = "total: {}(lines: {}), net: {}".format(total, lines, net)
            keymap.popBalloon("", t, 5000)
    keymap_global["LC-U1-C"] = lazy_call(count_chars)

    # wrap with quote mark
    def quote_selection() -> None:
        cb = copy_string()
        if cb:
            paste_string(' "{}" '.format(cb.strip()))
    keymap_global["LC-U0-Q"] = lazy_call(quote_selection)

    # paste with quote mark
    def paste_with_anchor(skip_blank:bool=False) -> callable:
        def _paster() -> None:
            cb = get_current_clipboard()
            lines = cb.strip().splitlines()
            quoted = []
            for line in lines:
                if skip_blank and len(line.strip()) < 1:
                    quoted.append(line)
                else:
                    quoted.append("> " + line)
            paste_string(os.linesep.join(quoted))
        return lazy_call(_paster)
    keymap_global["U1-Q"] = paste_with_anchor(False)
    keymap_global["C-U1-Q"] = paste_with_anchor(True)

    # open url in browser
    keymap_global["C-U0-O"] = lazy_call( lambda : execute_path(copy_string().strip()) )

    # re-input selected string with ime
    def re_input_with_ime() -> None:
        selection = copy_string()
        if selection:
            sequence = ["Minus" if c == "-" else c for c in to_non_space(selection)]
            set_ime(1)
            send_input(tuple(sequence), 0)
            delay()
    keymap_global["U1-I"] = lazy_call(recover_clipboard(re_input_with_ime))

    def moko(search_all:bool=False) -> callable:
        exe_path = resolve_path(r"Dropbox\develop\code\go\moko\src\moko.exe")
        def _launcher() -> None:
            execute_path(exe_path, "-src={} -filer={} -all={} -exclude=_obsolete".format(r"C:\Personal\launch.yaml", get_filer_path(), search_all))
        return lazy_call(_launcher)
    keymap_global["U1-Z"] = moko(False)
    keymap_global["LC-U1-Z"] = moko(True)



    ################################
    # config menu
    ################################

    def read_config() -> str:
        return Path(getAppExePath(), "config.py").read_text("utf-8")

    def open_github() -> None:
        repo = "https://github.com/AWtnb/keyhac/edit/main/config.py"
        setClipboardText(read_config())
        execute_path(repo)

    def reload_config() -> None:
        keymap.configure()
        keymap.updateKeymap()
        ts = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
        print("\n{} reloaded config.py\n".format(ts))

    keymap_global["LC-U0-X"] = keymap.defineMultiStrokeKeymap("config.py: R=>reload, E=>Edit, G=>Gist, P=>Paste")
    for key, func in {
        "R" : reload_config,
        "E" : keymap.command_EditConfig,
        "G" : open_github,
        "P" : lambda : paste_string(read_config()),
        "X" : lambda : None,
    }.items():
        keymap_global["LC-U0-X"][key] = lazy_call(func)


    ################################
    # class for position on monitor
    ################################

    class WndRect:
        def __init__(self, rect:list) -> None:
            self._rect = rect

            def _snapper() -> None:
                wnd = keymap.getTopLevelWindow()
                if list(wnd.getRect()) == self._rect:
                    wnd.maximize()
                else:
                    if wnd.isMaximized():
                        wnd.restore()
                        delay()
                trial_limit = 2
                while trial_limit > 0:
                    wnd.setRect(self._rect)
                    if list(wnd.getRect()) == self._rect:
                        return
                    trial_limit -= 1
            self.snap_func = _snapper


    class MonitorRect:
        def __init__(self, rect:list, idx:int) -> None:
            self.idx= idx
            self.left, self.top, self.right, self.bottom = rect
            self.max_width = self.right - self.left
            self.max_height = self.bottom - self.top

            def _size_dict(max:int) -> dict:
                return {
                    "small": int(max / 3),
                    "middle": int(max / 2),
                    "large": int(max * 2 / 3),
                }
            self.possible_width = _size_dict(self.max_width)
            self.possible_height = _size_dict(self.max_height)

            self.area_map = {}
            self._set_center_rect()
            self._set_horizontal_rect()
            self._set_vertical_rect()

        def _set_center_rect(self) -> None:
            d = {}
            for size,px in self.possible_width.items():
                lx = self.left + int((self.max_width - px) / 2)
                d[size] = WndRect([lx, self.top, lx + px, self.bottom])
            self.area_map["center"] = d

        def _set_horizontal_rect(self) -> None:
            for pos in ("left", "right"):
                d = {}
                for size,px in self.possible_width.items():
                    lx = self.left
                    if pos == "right":
                        lx = self.right - px
                    d[size] = WndRect([lx, self.top, lx + px, self.bottom])
                self.area_map[pos] = d

        def _set_vertical_rect(self) -> None:
            for pos in ("top", "bottom"):
                d = {}
                for size,px in self.possible_height.items():
                    ty = self.top
                    if pos == "bottom":
                        ty = self.bottom - px
                    d[size] = WndRect([self.left, ty, self.right, ty + px])
                self.area_map[pos] = d


    def get_monitors() -> list:
        monitors = [ MonitorRect(rect=m[1], idx=m[2]) for m in pyauto.Window.getMonitorInfo() ]
        max_widths = [m.max_width for m in monitors]
        if len(set(max_widths)) == 1:
            return sorted(monitors, key=lambda x : x.idx, reverse=True)
        return sorted(monitors, key=lambda x : x.max_width, reverse=True)

    KEYMAP_MONITORS = get_monitors()

    ################################
    # set cursor position
    ################################

    keymap_global["U0-Up"] = keymap.MouseWheelCommand(+0.5)
    keymap_global["U0-Down"] = keymap.MouseWheelCommand(-0.5)
    keymap_global["U0-Left"] = keymap.MouseHorizontalWheelCommand(-0.5)
    keymap_global["U0-Right"] = keymap.MouseHorizontalWheelCommand(+0.5)

    def set_cursor_pos(x:int, y:int) -> None:
        keymap.beginInput()
        keymap.input_seq.append(pyauto.MouseMove(x, y))
        keymap.endInput()

    def cursor_to_center() -> None:
        wnd = keymap.getTopLevelWindow()
        wnd_left, wnd_top, wnd_right, wnd_bottom = wnd.getRect()
        to_x = int((wnd_left + wnd_right) / 2)
        to_y = int((wnd_bottom + wnd_top) / 2)
        set_cursor_pos(to_x, to_y)
    keymap_global["O-(236)"] = cursor_to_center

    class CursorPos:
        def __init__(self) -> None:
            self.pos = []
            for monitor in KEYMAP_MONITORS:
                y = int(monitor.max_height / 2)
                for i in (1, 3):
                    x = monitor.left + int(monitor.max_width / 4) * i
                    self.pos.append([x, y])

        def get_snap_func(self) -> callable:
            xs = [p[0] for p in self.pos]
            def _snap() -> None:
                x, _ = pyauto.Input.getCursorPos()
                if x not in xs:
                    set_cursor_pos(*self.pos[0])
                else:
                    next_pos = self.pos[(xs.index(x)+1) % len(xs)]
                    set_cursor_pos(*next_pos)
            return _snap

    keymap_global["O-RCtrl"] = CursorPos().get_snap_func()



    ################################
    # set window position
    ################################

    keymap_global["U1-L"] = "LWin-Right"
    keymap_global["U1-H"] = "LWin-Left"

    keymap_global["U1-M"] = keymap.defineMultiStrokeKeymap("snap window position with key input!")
    keymap_global["U1-M"]["X"] = lazy_call(lambda : keymap.getTopLevelWindow().maximize())
    keymap_global["U1-M"]["F"] = "LWin-Home" #focus on current window

    for mod_mntr, mntr_idx in {
        "": 0,
        "A-": 1,
    }.items():
        for mod_area, size in {
            "": "middle",
            "S-": "large",
            "C-": "small",
        }.items():
            for key, pos in {
                "H": "left",
                "L": "right",
                "J": "bottom",
                "K": "top",
                "M": "center",
            }.items():
                if mntr_idx < len(KEYMAP_MONITORS):
                    snapper = KEYMAP_MONITORS[mntr_idx].area_map[pos][size].snap_func
                    keymap_global["U1-M"][mod_mntr+mod_area+key] = lazy_call(snapper)

    def snap_and_maximize(towards="Left") -> callable:
        def _snap() -> None:
            send_keys("LShift-LWin-"+towards)
            delay()
            keymap.getTopLevelWindow().maximize()
        return lazy_call(_snap)
    keymap_global["LC-U1-L"] = snap_and_maximize("Right")
    keymap_global["LC-U1-H"] = snap_and_maximize("Left")

    keymap_global["U1-M"]["U0-L"] = snap_and_maximize("Right")
    keymap_global["U1-M"]["U0-H"] = snap_and_maximize("Left")


    ################################
    # pseudo espanso
    ################################

    class PseudoEspanso:
        def __init__(self) -> None:
            self.mapping = {}

            direct_puncher = KeyPuncher(sleep_sec=0)
            for alias, stroke in {
                "-f": ("\uff0d"),
                "-h": ("\u2010"),
                "-m": ("\u2014"),
                "-n": ("\u2013"),
                "-s": ("\u2212"),
                "=": ("\u30a0"),
                "ae": ("\u00e9"),
                "aE": ("\u00c9"),
                "aee": ("\u00e8"),
                "aEE": ("\u00c8"),
                "md": ("div."),
                "mp": ("# ///"),
                "ms": ("span."),
                "np": ("proofed"),
                "npa": ("proofed_by_author"),
                "nsa": ("send_to_author"),
                "nsp": ("send_to_printshop"),
                "pa": (resolve_path(r"Dropbox\develop\app_config")),
                "pc": (resolve_path(r"Dropbox\develop\app_config\IME_google\convertion_dict")),
                "pcm": (resolve_path(r"Dropbox\develop\app_config\IME_google\convertion_dict\my.txt")),
                "pd": (resolve_path(r"Desktop")),
                "px": (resolve_path(r"Dropbox")),
                "rgbr": (r"[\[［].+?[\]］]"),
                "rgF": ("(?!)", "Left"),
                "rgf": ("(?=)", "Left"),
                "rgP": ("(?<!)", "Left"),
                "rgp": ("(?<=)", "Left"),
                "rgpr": (r"[\(（].+?[\)）]"),
                "uA": ("\u00c4"),
                "ua": ("\u00e4"),
                "uO": ("\u00d6"),
                "uo": ("\u00f6"),
                "uU": ("\u00dc"),
                "uu": ("\u00fc"),
                "xm": (".md"),
                "xx": (".txt"),
            }.items():
                self.mapping[alias] = direct_puncher.invoke(*stroke)

            indirect_puncher = KeyPuncher(recover_ime=True, sleep_sec=0)
            for alias, stroke in {
                "m1": ("# "),
                "m2": ("## "),
                "m3": ("### "),
                "m4": ("#### "),
                "m5": ("##### "),
                "m6": ("###### "),
            }.items():
                self.mapping[alias] = indirect_puncher.invoke(*stroke)

        def invoke(self) -> callable:
            def _sender() -> None:
                selection = copy_string()
                if (func := self.mapping.get(selection.strip())):
                    func()
                else:
                    send_keys("S-Left")
            return lazy_call(recover_clipboard(_sender))

    keymap_global["U1-X"] = PseudoEspanso().invoke()



    ################################
    # input customize
    ################################

    keymap_global["U1-W"] = keymap.defineMultiStrokeKeymap("Wrap with parenthesis:")

    # surround with brackets
    for key, pair in {
        "U0-2": ['"', '"'],
        "U0-7": ["'", "'"],
        "U0-AtMark": ["`", "`"],
        "U1-AtMark": [" `", "` "],
        "U0-CloseBracket": ["[", "]"],
        "U1-9": ["(", ")"],
        "U1-CloseBracket": ["{", "}"],
        "U0-T": ["<", "></>"],
        "U0-Caret": ["~~", "~~"],
    }.items():
        prefix, suffix = pair
        sent = pair + ["Left"]*len(suffix)
        keymap_global[key] = KeyPuncher().invoke(*sent)
        keymap_global["U1-W"][key] = KeyPuncher(delay_msec=50).invoke("C-Insert", prefix, "S-Insert", suffix)

    for key, pair in {
        "U0-8": ["\u300e", "\u300f"], # WHITE CORNER BRACKET 『』
        "U0-9": ["\u3010", "\u3011"], # BLACK LENTICULAR BRACKET 【】
        "U0-OpenBracket": ["\u300c", "\u300d"], # CORNER BRACKET 「」
        "U0-Y": ["\u300a", "\u300b"], # DOUBLE ANGLE BRACKET 《》
        "U1-2": ["\u201c", "\u201d"], # DOUBLE QUOTATION MARK “”
        "U1-7": ["\u3014", "\u3015"], # TORTOISE SHELL BRACKET 〔〕
        "U1-8": ["\uff08", "\uff09"], # FULLWIDTH PARENTHESIS （）
        "U1-OpenBracket": ["\uff3b", "\uff3d"], # FULLWIDTH SQUARE BRACKET ［］
        "U1-Y": ["\u3008", "\u3009"], # ANGLE BRACKET 〈〉
        "C-U0-Caret": ["\u300c", "\u300d\uff1f"],
    }.items():
        prefix, suffix = pair
        sent = pair + ["Left"]*len(suffix)
        keymap_global[key] = KeyPuncher(recover_ime=True).invoke(*sent)
        keymap_global["U1-W"][key] = KeyPuncher(recover_ime=True, delay_msec=50).invoke("C-Insert", prefix, "S-Insert", suffix)

    # input string without conversion even when ime is turned on
    def direct_input(key:str, turnoff_ime_later:bool=False) -> callable:
        finish_keys = ["C-M"]
        if turnoff_ime_later:
            finish_keys.append("(243)")
        def _input() -> None:
            send_keys(key)
            if keymap.getWindow().getImeStatus():
                send_keys(*finish_keys)
        return _input

    keymap_global["BackSlash"] = direct_input("S-BackSlash", False)

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
        keymap_global[key] = direct_input(key, turnoff_ime)

    for n in "123456789":
        key = "LS-" + n
        turnoff_ime = False
        if n in ("2", "3", "4"):
            turnoff_ime = True
        keymap_global[key] = direct_input(key, turnoff_ime)

    for alphabet in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        key = "LS-" + alphabet
        keymap_global[key] = direct_input(key, True)


    # input punctuation marks directly and set ime mode

    for key, seq in {
        "U0-4": ["S-4", "S-BackSlash"],
        "U0-Yen": ["S-Yen", "Left"],
        "U1-4": ["S-4", "S-BackSlash", "Period"],
    }.items():
        keymap_global[key] = KeyPuncher().invoke(*seq)

    for key, send in {
        "U0-1": "S-1",
        "U0-Colon": "Colon",
        "U0-Comma": "Comma",
        "U0-Period": "Period",
        "U0-Slash": "Slash",
        "U0-U": "S-BackSlash",
        "U1-Enter": "<br />",
        "U1-E": "S-Minus",
        "U1-Minus": "Minus",
    }.items():
        keymap_global[key] = KeyPuncher().invoke(send)

    for key, send in {
        "C-U0-P": "\uff01", # FULLWIDTH EXCLAMATION MARK
        "S-U0-Colon": "\uff1a", # FULLWIDTH COLON
        "S-U0-Comma": "\uff0c", # FULLWIDTH COMMA
        "S-U0-Minus": "\u3000\u2015\u2015",
        "S-U0-P": "\uff1f", # FULLWIDTH QUESTION MARK
        "S-U0-Period": "\uff0e", # FULLWIDTH FULL STOP
        "S-U0-U": "S-BackSlash",
        "U0-Minus": "\u2015\u2015", # HORIZONTAL BAR * 2
        "U0-P": "\u30fb", # KATAKANA MIDDLE DOT
        "U0-SemiColon": "+ ",
        "S-U0-8": "+ ",
        "U1-1": "1. ",
        "S-U0-7": "1. ",
    }.items():
        keymap_global[key] = KeyPuncher(recover_ime=True).invoke(send)


    # input and format date string
    def input_date(fmt:str, recover_ime:bool=False) -> callable:
        def _input() -> None:
            d = datetime.datetime.today()
            set_ime(0)
            send_input([c for c in d.strftime(fmt)], 0)
            if recover_ime:
                set_ime(1)
        return lazy_call(_input)

    keymap_global["U1-D"] = keymap.defineMultiStrokeKeymap("Delimiter: D(1)=>YMD, S(2)=>Y/M/D, P(3)=>Y.M.D, H(4)=>Y-M-D, J(5)=>Y年M月D日")
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
        keymap_global["U1-D"][key] = input_date(*params)


    ################################
    # web search
    ################################

    class SearchNoise:
        def __init__(self) -> None:
            self.space = ord(" ")
            self.kangxi_dict = { "\u2f00":"\u4e00", "\u2f01":"\u4e28", "\u2f02":"\u4e36", "\u2f03":"\u4e3f", "\u2f04":"\u4e59", "\u2f05":"\u4e85", "\u2f06":"\u4e8c", "\u2f07":"\u4ea0", "\u2f08":"\u4eba", "\u2f09":"\u513f", "\u2f0a":"\u5165", "\u2f0b":"\u516b", "\u2f0c":"\u5182", "\u2f0d":"\u5196", "\u2f0e":"\u51ab", "\u2f0f":"\u51e0", "\u2f10":"\u51f5", "\u2f11":"\u5200", "\u2f12":"\u529b", "\u2f13":"\u52f9", "\u2f14":"\u5315", "\u2f15":"\u531a", "\u2f16":"\u5338", "\u2f17":"\u5341", "\u2f18":"\u535c", "\u2f19":"\u5369", "\u2f1a":"\u5382", "\u2f1b":"\u53b6", "\u2f1c":"\u53c8", "\u2f1d":"\u53e3", "\u2f1e":"\u56d7", "\u2f1f":"\u571f", "\u2f20":"\u58eb", "\u2f21":"\u5902", "\u2f22":"\u590a", "\u2f23":"\u5915", "\u2f24":"\u5927", "\u2f25":"\u5973", "\u2f26":"\u5b50", "\u2f27":"\u5b80", "\u2f28":"\u5bf8", "\u2f29":"\u5c0f", "\u2f2a":"\u5c22", "\u2f2b":"\u5c38", "\u2f2c":"\u5c6e", "\u2f2d":"\u5c71", "\u2f2e":"\u5ddb", "\u2f2f":"\u5de5", "\u2f30":"\u5df1", "\u2f31":"\u5dfe", "\u2f32":"\u5e72", "\u2f33":"\u5e7a", "\u2f34":"\u5e7f", "\u2f35":"\u5ef4", "\u2f36":"\u5efe", "\u2f37":"\u5f0b", "\u2f38":"\u5f13", "\u2f39":"\u5f50", "\u2f3a":"\u5f61", "\u2f3b":"\u5f73", "\u2f3c":"\u5fc3", "\u2f3d":"\u6208", "\u2f3e":"\u6238", "\u2f3f":"\u624b", "\u2f40":"\u652f", "\u2f41":"\u6534", "\u2f42":"\u6587", "\u2f43":"\u6597", "\u2f44":"\u65a4", "\u2f45":"\u65b9", "\u2f46":"\u65e0", "\u2f47":"\u65e5", "\u2f48":"\u66f0", "\u2f49":"\u6708", "\u2f4a":"\u6728", "\u2f4b":"\u6b20", "\u2f4c":"\u6b62", "\u2f4d":"\u6b79", "\u2f4e":"\u6bb3", "\u2f4f":"\u6bcb", "\u2f50":"\u6bd4", "\u2f51":"\u6bdb", "\u2f52":"\u6c0f", "\u2f53":"\u6c14", "\u2f54":"\u6c34", "\u2f55":"\u706b", "\u2f56":"\u722a", "\u2f57":"\u7236", "\u2f58":"\u723b", "\u2f59":"\u723f", "\u2f5a":"\u7247", "\u2f5b":"\u7259", "\u2f5c":"\u725b", "\u2f5d":"\u72ac", "\u2f5e":"\u7384", "\u2f5f":"\u7389", "\u2f60":"\u74dc", "\u2f61":"\u74e6", "\u2f62":"\u7518", "\u2f63":"\u751f", "\u2f64":"\u7528", "\u2f65":"\u7530", "\u2f66":"\u758b", "\u2f67":"\u7592", "\u2f68":"\u7676", "\u2f69":"\u767d", "\u2f6a":"\u76ae", "\u2f6b":"\u76bf", "\u2f6c":"\u76ee", "\u2f6d":"\u77db", "\u2f6e":"\u77e2", "\u2f6f":"\u77f3", "\u2f70":"\u793a", "\u2f71":"\u79b8", "\u2f72":"\u79be", "\u2f73":"\u7a74", "\u2f74":"\u7acb", "\u2f75":"\u7af9", "\u2f76":"\u7c73", "\u2f77":"\u7cf8", "\u2f78":"\u7f36", "\u2f79":"\u7f51", "\u2f7a":"\u7f8a", "\u2f7b":"\u7fbd", "\u2f7c":"\u8001", "\u2f7d":"\u800c", "\u2f7e":"\u8012", "\u2f7f":"\u8033", "\u2f80":"\u807f", "\u2f81":"\u8089", "\u2f82":"\u81e3", "\u2f83":"\u81ea", "\u2f84":"\u81f3", "\u2f85":"\u81fc", "\u2f86":"\u820c", "\u2f87":"\u821b", "\u2f88":"\u821f", "\u2f89":"\u826e", "\u2f8a":"\u8272", "\u2f8b":"\u8278", "\u2f8c":"\u864d", "\u2f8d":"\u866b", "\u2f8e":"\u8840", "\u2f8f":"\u884c", "\u2f90":"\u8863", "\u2f91":"\u897e", "\u2f92":"\u898b", "\u2f93":"\u89d2", "\u2f94":"\u8a00", "\u2f95":"\u8c37", "\u2f96":"\u8c46", "\u2f97":"\u8c55", "\u2f98":"\u8c78", "\u2f99":"\u8c9d", "\u2f9a":"\u8d64", "\u2f9b":"\u8d70", "\u2f9c":"\u8db3", "\u2f9d":"\u8eab", "\u2f9e":"\u8eca", "\u2f9f":"\u8f9b", "\u2fa0":"\u8fb0", "\u2fa1":"\u8fb5", "\u2fa2":"\u9091", "\u2fa3":"\u9149", "\u2fa4":"\u91c6", "\u2fa5":"\u91cc", "\u2fa6":"\u91d1", "\u2fa7":"\u9577", "\u2fa8":"\u9580", "\u2fa9":"\u961c", "\u2faa":"\u96b6", "\u2fab":"\u96b9", "\u2fac":"\u96e8", "\u2fad":"\u9751", "\u2fae":"\u975e", "\u2faf":"\u9762", "\u2fb0":"\u9769", "\u2fb1":"\u97cb", "\u2fb2":"\u97ed", "\u2fb3":"\u97f3", "\u2fb4":"\u9801", "\u2fb5":"\u98a8", "\u2fb6":"\u98db", "\u2fb7":"\u98df", "\u2fb8":"\u9996", "\u2fb9":"\u9999", "\u2fba":"\u99ac", "\u2fbb":"\u9aa8", "\u2fbc":"\u9ad8", "\u2fbd":"\u9adf", "\u2fbe":"\u9b25", "\u2fbf":"\u9b2f", "\u2fc0":"\u9b32", "\u2fc1":"\u9b3c", "\u2fc2":"\u9b5a", "\u2fc3":"\u9ce5", "\u2fc4":"\u9e75", "\u2fc5":"\u9e7f", "\u2fc6":"\u9ea5", "\u2fc7":"\u9ebb", "\u2fc8":"\u9ec3", "\u2fc9":"\u9ecd", "\u2fca":"\u9ed2", "\u2fcb":"\u9ef9", "\u2fcc":"\u9efd", "\u2fcd":"\u9f0e", "\u2fce":"\u9f13", "\u2fcf":"\u9f20", "\u2fd0":"\u9f3b", "\u2fd1":"\u9f4a", "\u2fd2":"\u9f52", "\u2fd3":"\u9f8d", "\u2fd4":"\u9f9c", "\u2fd5":"\u9fa0" }

            self.punc_dict = { int("30FB",16): self.space }
            for noise_range in [
                ["0021", "002F"],
                ["003A", "0040"],
                ["005B", "0060"],
                ["007B", "007E"],
                # quotation marks
                ["2018", "201F"],
                # horizontal bars
                ["2010", "2017"],
                ["2500", "2501"],
                ["2E3A", "2E3B"],
                # fullwidth symbols
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
                # CJK Radicals
                ["2E80", "2EF3"],
            ]:
                f, t = noise_range
                for i in range(int(f,16), int(t,16)+1):
                    self.punc_dict[i] = self.space

            self.hiragana_dict = {}
            for i in range(int("3041",16), int("3093",16)+1):
                self.hiragana_dict[i] = self.space

        def to_space(self, s:str, strip_hiraganas:bool) -> str:
            fmt = s.translate(str.maketrans(self.punc_dict)).translate(str.maketrans(self.kangxi_dict))
            if not strip_hiraganas:
                return fmt
            return fmt.translate(str.maketrans(self.hiragana_dict))


    SEARCH_NOISE = SearchNoise()

    class SearchQuery:
        def __init__(self, query:str, strip_hiragana:bool=False) -> None:
            self._query = ""
            self._set_query(query)
            self._sanitize_query(strip_hiragana)

        def _set_query(self, s:str) -> str:
            lines = s.strip().replace("\u200b", "").replace("　", " ").replace("\t", " ").splitlines()
            self._query = "".join([ line.rstrip("-") for line in lines ])

        def _sanitize_query(self, strip_hiragana:bool) -> None:
            for honor in ["先生", "様", "監修", "et al."]:
                self._query = self._query.replace(honor, " ")
            self._query = SEARCH_NOISE.to_space(self._query, strip_hiragana)

        def _get_words(self) -> list:
            words = self._query.split(" ")
            return [w for w in words if w]

        def quote_each_word(self) -> None:
            words = self._get_words()
            self._query = " ".join(['"{}"'.format(w) for w in words])

        def encode(self) -> str:
            return urllib.parse.quote(" ".join(self._get_words()))

    def search_on_web(uri:str, quote_each:bool=False, strip_hiragana:bool=False) -> callable:
        def _search() -> None:
            s = copy_string()
            q = SearchQuery(s, strip_hiragana)
            if quote_each:
                q.quote_each_word()
            execute_path(uri.format(q.encode()))
        return lazy_call(_search)

    for mdf, params in {
        "": (False, False),
        "S-": (True, False),
        "C-": (False, True),
        "S-C-": (True, True),
    }.items():
        keymap_global[mdf+"U0-S"] = keymap.defineMultiStrokeKeymap("quote-each:{} / strip-hiragana:{}".format(*params))
        for key, uri in {
            "A": "https://www.amazon.co.jp/s?i=stripbooks&k={}",
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
            "S": "https://scholar.google.co.jp/scholar?q={}",
            "T": "https://twitter.com/search?q={}",
            "Y": "http://www.google.co.jp/search?q=site%3Ayuhikaku.co.jp%20{}",
            "W": "https://www.worldcat.org/search?q={}",
        }.items():
            keymap_global[mdf+"U0-S"][key] = search_on_web(uri, *params)

    ################################
    # activate window
    ################################

    class SystemBrowser:
        def __init__(self) -> None:
            self.prog_id = self._get_prog_id()
            self.commandline = self._get_commandline()
            self.exe_path = self._get_exe_path()
            self.exe_name = self._get_exe_name()
            self.wnd_class = self._get_wnd_class()

        def _get_prog_id(self) -> str:
            register_path = r'Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice'
            with OpenKey(HKEY_CURRENT_USER, register_path) as key:
                return str(QueryValueEx(key, "ProgId")[0])

        def _get_commandline(self) -> str:
            register_path = r'{}\shell\open\command'.format(self.prog_id)
            with OpenKey(HKEY_CLASSES_ROOT, register_path) as key:
                return str(QueryValueEx(key, "")[0])

        def _get_exe_path(self) -> str:
            return re.sub(r"(^.+\.exe)(.*)", r"\1", self.commandline).replace('"', "")

        def _get_exe_name(self) -> str:
            return Path(self._get_exe_path()).name

        def _get_wnd_class(self) -> str:
            return {
                "chrome.exe": "Chrome_WidgetWin_1",
                "vivaldi.exe": "Chrome_WidgetWin_1",
                "firefox.exe": "MozillaWindowClass",
            }.get(self._get_exe_name(), "Chrome_WidgetWin_1")

    DEFAULT_BROWSER = SystemBrowser()


    class PyWnd:
        def __init__(self, exe_name:str, class_name:str) -> None:
            self.exe_name = exe_name
            self.class_name = class_name
            self.target = None
            self._find()

        def _find(self) -> pyauto.Window:
            # https://github.com/crftwr/keyhac/blob/master/keyhac_keymap.py
            def __hop_to_next(wnd:pyauto.Window, arg) -> bool:
                if not wnd.isVisible() : return True
                if not fnmatch.fnmatch(wnd.getProcessName(), self.exe_name) : return True
                if self.class_name and not fnmatch.fnmatch(wnd.getClassName(), self.class_name) : return True
                self.target = wnd.getLastActivePopup()
                return False
            pyauto.Window.enum(__hop_to_next, None)

        def activate(self) -> bool:
            if not self.target:
                return False
            if self.target.isMinimized():
                self.target.restore()
            timeout = 0.25
            retry_time = 0.01
            while timeout > 0.0:
                try:
                    self.target.setForeground()
                    if pyauto.Window.getForeground() == self.target:
                        delay()
                        self.target.setForeground(True)
                        return True
                except:
                    return False
                delay(retry_time)
                timeout -= retry_time
            return False



    def pseudo_cuteExec(exe_name:str, class_name:str, exe_path:str) -> callable:
        def _executer() -> None:
            w = PyWnd(exe_name, class_name)
            if w.target:
                if keymap.getWindow() == w.target or not w.activate():
                    send_keys("LCtrl-LAlt-Tab")
            else:
                execute_path(exe_path)
        return lazy_call(_executer, 0)

    keymap_global["U1-C"] = keymap.defineMultiStrokeKeymap()
    for key, params in {
        "Space": (
            DEFAULT_BROWSER.exe_name,
            DEFAULT_BROWSER.wnd_class,
            DEFAULT_BROWSER.exe_path
        ),
        "C": (
            "chrome.exe",
            "Chrome_WidgetWin_1",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ),
        "S": (
            "slack.exe",
            "Chrome_WidgetWin_1",
            resolve_path(r"AppData\Local\slack\slack.exe")
        ),
        "F": (
            "firefox.exe",
            "MozillaWindowClass",
            r"C:\Program Files\Mozilla Firefox\firefox.exe"
        ),
        "B": (
            "thunderbird.exe",
            "MozillaWindowClass",
            r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"
        ),
        "O": (
            "Obsidian.exe",
            "Chrome_WidgetWin_1",
            None
        ),
        "P": (
            "SumatraPDF.exe",
            "SUMATRA_PDF_FRAME",
            None
        ),
        "C-P": (
            "powerpnt.exe",
            "PPTFrameClass",
            None
        ),
        "E": (
            "EXCEL.EXE",
            "XLMAIN",
            None
        ),
        "W": (
            "WINWORD.EXE",
            "OpusApp",
            None
        ),
        "V": (
            "Code.exe",
            "Chrome_WidgetWin_1",
            resolve_path(r"scoop\apps\vscode\current\Code.exe")
        ),
        "C-V": (
            "vivaldi.exe",
            "Chrome_WidgetWin_1",
            None
        ),
        "T": (
            "TE64.exe",
            "TablacusExplorer",
            resolve_path(r"Dropbox\portable_apps\tablacus\TE64.exe")
        ),
        "M": (
            "Mery.exe",
            "TChildForm",
            resolve_path(r"AppData\Local\Programs\Mery\Mery.exe")
        ),
        "X": (
            "explorer.exe",
            "CabinetWClass",
            r"C:\Windows\explorer.exe"
        )
    }.items():
        keymap_global["U1-C"][key] = pseudo_cuteExec(*params)


    # activate application
    for key, params in {
        "U1-T": (
            "TE64.exe",
            "TablacusExplorer",
            None
        ),
        "U1-P": (
            "SumatraPDF.exe",
            "SUMATRA_PDF_FRAME",
            None
        ),
        "LC-U1-M": (
            "Mery.exe",
            "TChildForm",
            resolve_path(r"AppData\Local\Programs\Mery\Mery.exe")
        ),
        "LC-U1-N": (
            "notepad.exe",
            "Notepad",
            r"C:\Windows\System32\notepad.exe"
        ),
        "C-U1-W": (
            "WindowsTerminal.exe",
            "CASCADIA_HOSTING_WINDOW_CLASS",
            None
        )
    }.items():
        keymap_global[key] = pseudo_cuteExec(*params)

    keymap_global["U1-M"]["U1-M"] = pseudo_cuteExec("Mery.exe", "TChildForm", None)

    # def check_window_class() -> None:
    #     print("process name:")
    #     print("    - {}".format(keymap.getWindow().getProcessName()))
    #     print("window class:")
    #     print("    - current: {}".format(keymap.getWindow().getClassName()))
    #     print("    - toplevel: {}".format(keymap.getTopLevelWindow().getClassName()))
    #     print("window text:")
    #     print("    - {}".format(keymap.getWindow().getText() or ""))
    # keymap_global["LC-U1-F1"] = check_window_class

    # invoke specific filer
    def invoke_filer(dir_path:str) -> callable:
        filer_path = get_filer_path()
        if test_path(dir_path):
            def _invoker() -> None:
                execute_path(filer_path, dir_path)
            return lazy_call(_invoker)
        else:
            return lambda : None
    keymap_global["U1-F"] = keymap.defineMultiStrokeKeymap("invoke directory on filer:")
    keymap_global["U1-F"]["D"] = invoke_filer(resolve_path(r"Desktop"))
    keymap_global["U1-F"]["S"] = invoke_filer(resolve_path(r"Desktop\scan"))

    def invoke_cmder() -> None:
        wnd = PyWnd("ConEmu64.exe", "VirtualConsoleClass")
        if wnd.target:
            send_keys("C-AtMark")
        else:
            cmder_path = resolve_path(r"scoop\apps\cmder\current\Cmder.exe")
            execute_path(cmder_path)
    keymap_global["LC-AtMark"] = lazy_call(invoke_cmder)

    def search_on_browser() -> callable:
        def _invoker() -> None:
            if keymap.getWindow().getProcessName() == DEFAULT_BROWSER.exe_name:
                send_keys("C-T")
            else:
                w = PyWnd(DEFAULT_BROWSER.exe_name, DEFAULT_BROWSER.wnd_class)
                if w.target:
                    if w.activate():
                        delay()
                        send_keys("C-T")
                    else:
                        send_keys("LCtrl-LAlt-Tab")
                else:
                    execute_path("https://duckduckgo.com")
        return lazy_call(_invoker, 100)
    keymap_global["U0-Q"] =search_on_browser()

    ################################
    # application based remap
    ################################

    # browser
    keymap_browser = keymap.defineWindowKeymap(check_func=is_browser_wnd)
    keymap_browser["LC-LS-W"] = "A-Left"
    keymap_browser["O-LShift"] = "C-F"
    keymap_browser["LC-Q"] = "A-F4"

    # intra
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClient.exe")
    keymap_intra["O-(235)"] = lambda : None

    # slack
    keymap_slack = keymap.defineWindowKeymap(exe_name="slack.exe", class_name="Chrome_WidgetWin_1")
    keymap_slack["O-LShift"] = "C-K"
    keymap_slack["F3"] = "C-K"
    keymap_slack["F1"] = KeyPuncher().invoke("+:")

    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")
    keymap_vscode["C-S-P"] = KeyPuncher().invoke("C-S-P")

    # mery
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")
    keymap_mery["LA-U0-J"] = "A-CloseBracket"
    keymap_mery["LA-U0-K"] = "A-OpenBracket"
    keymap_mery["LA-LS-U0-J"] = "A-S-CloseBracket"
    keymap_mery["LA-LS-U0-k"] = "A-S-OpenBracket"

    # cmder
    keymap_cmder = keymap.defineWindowKeymap(class_name="VirtualConsoleClass")
    keymap_cmder["LAlt-Space"] = "Lwin-LAlt-Space"
    keymap_cmder["C-W"] = lambda : None

    # sumatra PDF
    keymap_sumatra = keymap.defineWindowKeymap(exe_name="SumatraPDF.exe")
    keymap_sumatra["C-B"] = "F12"
    keymap_sumatra["C-K"] = "S-F3"
    keymap_sumatra["C-J"] = "F3"
    keymap_sumatra["C-OpenBracket"] = "S-F3"
    keymap_sumatra["C-CloseBracket"] = "F3"

    keymap_sumatra["O-LShift"] = KeyPuncher(recover_ime=True, delay_msec=0).invoke("Esc", "C-Home", "C-F")

    # sumatra PDF when focus out from textbox
    keymap_sumatra_view = keymap.defineWindowKeymap(check_func=lambda wnd : wnd.getProcessName() == "SumatraPDF.exe" and wnd.getClassName() != "Edit")

    for key, seq in {
        "F": ["C-F"],
        "G": ["C-G"],
        "C": ["C"],
        "Z": ["Z"],
        "J": ["J"],
        "K": ["K"],
        "Q": ["Q"],
        "Enter": ["F3"],
        "S-Enter": ["S-F3"],
        "R": ["C-S-SemiColon"],
        "L": ["C-Tab"],
        "H": ["C-S-Tab"],
    }.items():
        # seq = ["Esc"] + seq
        keymap_sumatra_view[key] = KeyPuncher(sleep_sec=0.05).invoke(*seq)

    # word
    keymap_word = keymap.defineWindowKeymap(exe_name="WINWORD.EXE")
    keymap_word["F11"] = KeyPuncher().invoke("A-T", "P", "A-N")
    keymap_word["C-G"] = KeyPuncher().invoke("C-G")
    keymap_word["LC-Q"] = "A-F4"
    keymap_word["O-LShift"] = "C-F"

    # powerpoint
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    keymap_ppt["F11"] = KeyPuncher().invoke("A-T", "Down", "Home")

    # excel
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")
    keymap_excel["U0-M"] = "Enter", "Enter"
    keymap_excel["F11"] = KeyPuncher().invoke("A-T", "P", "A-N")
    keymap_excel["LC-Q"] = "A-F4"
    keymap_excel["O-LShift"] = "C-F"

    def select_all() -> None:
        if keymap.getWindow().getClassName() == "EXCEL6":
            send_keys("C-End", "C-S-Home")
        else:
            send_keys("C-A")
    keymap_excel["C-A"] = select_all

    def select_cell_content() -> None:
        if keymap.getWindow().getClassName() == "EXCEL7":
            send_keys("F2", "C-S-Home")
    keymap_excel["LC-U0-N"] = select_cell_content

    # Thunderbird
    keymap_tb = keymap.defineWindowKeymap(exe_name="thunderbird.exe")
    keymap_tb["C-S-V"] = "A-S", "Tab", "C-V", "C-Home", "S-End"
    keymap_tb["C-S-S"] = "C-X", "Delete", "A-S", "C-V"

    # filer
    keymap_filer = keymap.defineWindowKeymap(check_func=lambda wnd : wnd.getProcessName() in ("explorer.exe", "TE64.exe"))
    keymap_filer["C-S-C"] = "C-Add"
    keymap_filer["C-L"] = "A-D", "C-C"
    keymap_filer["A-K"] = "A-Up"

    ################################
    # popup clipboard menu
    ################################

    # enclosing functions for pop-up menu
    def format_cb(func:callable) -> callable:
        def _formatter() -> str:
            cb = get_current_clipboard()
            if cb:
                return func(cb)
        return _formatter
    def replace_cb(search:str, replace_to:str) -> callable:
        reg = re.compile(search)
        def _replacer() -> str:
            cb = get_current_clipboard()
            if cb:
                return reg.sub(replace_to, cb)
        return _replacer

    def catanate_file_content(s:str) -> str:
        if test_path(s):
            return Path(s).read_text("utf-8")
        return None

    def skip_blank_line(s:str) -> str:
        lines = s.strip().splitlines()
        return os.linesep.join([l for l in lines if l.strip()])

    class CharWidth:
        def __init__(self, focus_punctuation:bool=False) -> None:
            if focus_punctuation:
                self.full_letters = "\uff08\uff09\uff3b\uff3d"
                self.half_letters = "()[]"
            else:
                self.full_letters = "\uff41\uff42\uff43\uff44\uff45\uff46\uff47\uff48\uff49\uff4a\uff4b\uff4c\uff4d\uff4e\uff4f\uff50\uff51\uff52\uff53\uff54\uff55\uff56\uff57\uff58\uff59\uff5a\uff21\uff22\uff23\uff24\uff25\uff26\uff27\uff28\uff29\uff2a\uff2b\uff2c\uff2d\uff2e\uff2f\uff30\uff31\uff32\uff33\uff34\uff35\uff36\uff37\uff38\uff39\uff3a\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19\uff0d"
                self.half_letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
        def to_half_width(self, s:str) -> str:
            return s.translate(str.maketrans(self.full_letters, self.half_letters))
        def to_full_width(self, s:str) -> str:
            return s.translate(str.maketrans(self.half_letters, self.full_letters))

    class TitleCase:
        def __init__(self, s:str) -> None:
            self.text = s[0].upper() + s[1:]
            self.fix_space()
            self.fix_colon()
        def fix_space(self) -> None:
            self.text = re.sub(r"\s+", " ", self.text)
        def fix_colon(self) -> None:
            s = self.text.replace(" :", ":")
            self.text = re.sub(r":(?! )", ": ", s)
        def format(self) -> str:
            reg = re.compile(r"[- ][a-z]")
            def _replacer(mo:re.Match) -> str:
                return mo.group(0).upper()
            return reg.sub(_replacer, self.text)

    def to_smart_title_case(s:str) -> str:
        title_cased = TitleCase(s).format()
        words = title_cased.split(" ")
        stack = [words[0]]
        lower_list = ["A", "About", "Aboard", "Above", "Across", "After", "Against", "Along", "Alongside", "Amid", "Among", "An", "And", "Anti", "Around", "As", "At", "Bar", "Before", "Behind", "Below", "Beneath", "Beside", "Besides", "Between", "Beyond", "But", "By", "Considering", "Despite", "Down", "During", "Except", "For", "From", "In", "Inside", "Into", "Less", "Like", "Minus", "Near", "Notwithstanding", "Of", "Off", "On", "Onto", "Or", "Opposite", "Out", "Outside", "Over", "Pace", "Past", "Pending", "Per", "Plus", "Re", "Regarding", "Round", "Save", "Saving", "Since", "Than", "The", "Through", "Throughout", "Till", "To", "Touching", "Toward", "Under", "Underneath", "Unless", "Unlike", "Until", "Up", "Versus", "Via", "Vice", "With", "Within", "Without"]
        for word in words[1:]:
            if (word in lower_list) and (not stack[-1].endswith(":")):
                stack.append(word.lower())
            else:
                stack.append(word)
        return " ".join(stack)

    def format_kangxi_radicals(s:str) -> str:
        table = { "\u2f00":"\u4e00", "\u2f01":"\u4e28", "\u2f02":"\u4e36", "\u2f03":"\u4e3f", "\u2f04":"\u4e59", "\u2f05":"\u4e85", "\u2f06":"\u4e8c", "\u2f07":"\u4ea0", "\u2f08":"\u4eba", "\u2f09":"\u513f", "\u2f0a":"\u5165", "\u2f0b":"\u516b", "\u2f0c":"\u5182", "\u2f0d":"\u5196", "\u2f0e":"\u51ab", "\u2f0f":"\u51e0", "\u2f10":"\u51f5", "\u2f11":"\u5200", "\u2f12":"\u529b", "\u2f13":"\u52f9", "\u2f14":"\u5315", "\u2f15":"\u531a", "\u2f16":"\u5338", "\u2f17":"\u5341", "\u2f18":"\u535c", "\u2f19":"\u5369", "\u2f1a":"\u5382", "\u2f1b":"\u53b6", "\u2f1c":"\u53c8", "\u2f1d":"\u53e3", "\u2f1e":"\u56d7", "\u2f1f":"\u571f", "\u2f20":"\u58eb", "\u2f21":"\u5902", "\u2f22":"\u590a", "\u2f23":"\u5915", "\u2f24":"\u5927", "\u2f25":"\u5973", "\u2f26":"\u5b50", "\u2f27":"\u5b80", "\u2f28":"\u5bf8", "\u2f29":"\u5c0f", "\u2f2a":"\u5c22", "\u2f2b":"\u5c38", "\u2f2c":"\u5c6e", "\u2f2d":"\u5c71", "\u2f2e":"\u5ddb", "\u2f2f":"\u5de5", "\u2f30":"\u5df1", "\u2f31":"\u5dfe", "\u2f32":"\u5e72", "\u2f33":"\u5e7a", "\u2f34":"\u5e7f", "\u2f35":"\u5ef4", "\u2f36":"\u5efe", "\u2f37":"\u5f0b", "\u2f38":"\u5f13", "\u2f39":"\u5f50", "\u2f3a":"\u5f61", "\u2f3b":"\u5f73", "\u2f3c":"\u5fc3", "\u2f3d":"\u6208", "\u2f3e":"\u6238", "\u2f3f":"\u624b", "\u2f40":"\u652f", "\u2f41":"\u6534", "\u2f42":"\u6587", "\u2f43":"\u6597", "\u2f44":"\u65a4", "\u2f45":"\u65b9", "\u2f46":"\u65e0", "\u2f47":"\u65e5", "\u2f48":"\u66f0", "\u2f49":"\u6708", "\u2f4a":"\u6728", "\u2f4b":"\u6b20", "\u2f4c":"\u6b62", "\u2f4d":"\u6b79", "\u2f4e":"\u6bb3", "\u2f4f":"\u6bcb", "\u2f50":"\u6bd4", "\u2f51":"\u6bdb", "\u2f52":"\u6c0f", "\u2f53":"\u6c14", "\u2f54":"\u6c34", "\u2f55":"\u706b", "\u2f56":"\u722a", "\u2f57":"\u7236", "\u2f58":"\u723b", "\u2f59":"\u723f", "\u2f5a":"\u7247", "\u2f5b":"\u7259", "\u2f5c":"\u725b", "\u2f5d":"\u72ac", "\u2f5e":"\u7384", "\u2f5f":"\u7389", "\u2f60":"\u74dc", "\u2f61":"\u74e6", "\u2f62":"\u7518", "\u2f63":"\u751f", "\u2f64":"\u7528", "\u2f65":"\u7530", "\u2f66":"\u758b", "\u2f67":"\u7592", "\u2f68":"\u7676", "\u2f69":"\u767d", "\u2f6a":"\u76ae", "\u2f6b":"\u76bf", "\u2f6c":"\u76ee", "\u2f6d":"\u77db", "\u2f6e":"\u77e2", "\u2f6f":"\u77f3", "\u2f70":"\u793a", "\u2f71":"\u79b8", "\u2f72":"\u79be", "\u2f73":"\u7a74", "\u2f74":"\u7acb", "\u2f75":"\u7af9", "\u2f76":"\u7c73", "\u2f77":"\u7cf8", "\u2f78":"\u7f36", "\u2f79":"\u7f51", "\u2f7a":"\u7f8a", "\u2f7b":"\u7fbd", "\u2f7c":"\u8001", "\u2f7d":"\u800c", "\u2f7e":"\u8012", "\u2f7f":"\u8033", "\u2f80":"\u807f", "\u2f81":"\u8089", "\u2f82":"\u81e3", "\u2f83":"\u81ea", "\u2f84":"\u81f3", "\u2f85":"\u81fc", "\u2f86":"\u820c", "\u2f87":"\u821b", "\u2f88":"\u821f", "\u2f89":"\u826e", "\u2f8a":"\u8272", "\u2f8b":"\u8278", "\u2f8c":"\u864d", "\u2f8d":"\u866b", "\u2f8e":"\u8840", "\u2f8f":"\u884c", "\u2f90":"\u8863", "\u2f91":"\u897e", "\u2f92":"\u898b", "\u2f93":"\u89d2", "\u2f94":"\u8a00", "\u2f95":"\u8c37", "\u2f96":"\u8c46", "\u2f97":"\u8c55", "\u2f98":"\u8c78", "\u2f99":"\u8c9d", "\u2f9a":"\u8d64", "\u2f9b":"\u8d70", "\u2f9c":"\u8db3", "\u2f9d":"\u8eab", "\u2f9e":"\u8eca", "\u2f9f":"\u8f9b", "\u2fa0":"\u8fb0", "\u2fa1":"\u8fb5", "\u2fa2":"\u9091", "\u2fa3":"\u9149", "\u2fa4":"\u91c6", "\u2fa5":"\u91cc", "\u2fa6":"\u91d1", "\u2fa7":"\u9577", "\u2fa8":"\u9580", "\u2fa9":"\u961c", "\u2faa":"\u96b6", "\u2fab":"\u96b9", "\u2fac":"\u96e8", "\u2fad":"\u9751", "\u2fae":"\u975e", "\u2faf":"\u9762", "\u2fb0":"\u9769", "\u2fb1":"\u97cb", "\u2fb2":"\u97ed", "\u2fb3":"\u97f3", "\u2fb4":"\u9801", "\u2fb5":"\u98a8", "\u2fb6":"\u98db", "\u2fb7":"\u98df", "\u2fb8":"\u9996", "\u2fb9":"\u9999", "\u2fba":"\u99ac", "\u2fbb":"\u9aa8", "\u2fbc":"\u9ad8", "\u2fbd":"\u9adf", "\u2fbe":"\u9b25", "\u2fbf":"\u9b2f", "\u2fc0":"\u9b32", "\u2fc1":"\u9b3c", "\u2fc2":"\u9b5a", "\u2fc3":"\u9ce5", "\u2fc4":"\u9e75", "\u2fc5":"\u9e7f", "\u2fc6":"\u9ea5", "\u2fc7":"\u9ebb", "\u2fc8":"\u9ec3", "\u2fc9":"\u9ecd", "\u2fca":"\u9ed2", "\u2fcb":"\u9ef9", "\u2fcc":"\u9efd", "\u2fcd":"\u9f0e", "\u2fce":"\u9f13", "\u2fcf":"\u9f20", "\u2fd0":"\u9f3b", "\u2fd1":"\u9f4a", "\u2fd2":"\u9f52", "\u2fd3":"\u9f8d", "\u2fd4":"\u9f9c", "\u2fd5":"\u9fa0" }
        return s.translate(str.maketrans(table))

    def strip_postalcode_from_address(s:str) -> str:
        reg = re.compile(r"(\d{3}).(\d{4})[\s\r\n]*(.+$)")
        hankaku = CharWidth().to_half_width(s.strip().strip("\u3012"))
        m = reg.match(hankaku)
        if m:
            return "{}-{}\t{}".format(m.group(1), m.group(2), m.group(3))
        return s

    def fix_dumb_quotation(s:str) -> str:
        reg = re.compile(r"\"([^\"]+?)\"|'([^']+?)'")
        def _replacer(mo:re.Match) -> str:
            if str(mo.group(0)).startswith('"'):
                return "\u201c{}\u201d".format(mo.group(1))
            return "\u2018{}\u2019".format(mo.group(1))
        return reg.sub(_replacer, s)

    def decode_url(s:str) -> str:
        return urllib.parse.unquote(s)

    def to_double(s:str) -> str:
        symbol_map = {
            "\u300c": "\u300e",
            "\u300d": "\u300f",
            "\u2018": "\u201c",
            "\u2019": "\u201d",
            "'": '"',
        }
        # target: 「」‘’'
        reg = re.compile(r"[\u300c\u300d\u2018\u2019\u0027]")
        def _replacer(mo:re.Match) -> str:
            return symbol_map[mo.group(0)]
        return reg.sub(_replacer, s)

    def to_single(s:str) -> str:
        symbol_map = {
            "\u300e": "\u300c",
            "\u300f": "\u300d",
            "\u201c": "\u2018",
            "\u201d": "\u2019",
            '"': "'",
        }
        # target: 『』“”"
        reg = re.compile(r'[\u300e\u300f\u201c\u201d\u0022]')
        def _replacer(mo:re.Match) -> str:
            return symbol_map[mo.group(0)]
        return reg.sub(_replacer, s)

    def to_tortoise(s:str) -> str:
        reg = re.compile(r"[\u0028\u0029\uff08\uff09]")
        def _replacer(mo:re.Match) -> str:
            if mo.group(0) in ("\uff08", "\u0028"):
                return "\u3014"
            return "\u3015"
        return reg.sub(_replacer, s)

    def format_zoom_invitation(s:str) -> str:
        def _format_time(mo:re.Match) -> str:
            d = mo.group(1).strip()
            ymd = [int(s) for s in  re.split("[年月日]", d.strip("時間:：日"))]
            week = "（{}）".format("月火水木金土日"[datetime.date(*ymd).weekday()])
            t = mo.group(2)
            if t.endswith("PM"):
                h = int(t[:2]) + 12
                m = t[3:5]
                return "{d}{week} {h}:{m}～".format(d=d, week=week, h=h, m=m)
            return d + week + " AM " + t[:5] + "～"
        lines = s.replace(": ", "\uff1a").strip().splitlines()
        if len(lines) < 9:
            return None
        due = re.sub(r"^(.+日 )(.+[AP]M).+$", _format_time, lines[3])
        return os.linesep.join([
            "------------------------------",
            lines[2],
            due,
            lines[6],
            lines[8],
            lines[9],
            "------------------------------",
        ])

    keymap_global[ "LC-LS-X" ] = keymap.command_ClipboardList

    for title, menu in {
        "Noise-Reduction": [
            (" Remove blank line ", format_cb(skip_blank_line) ),
            (" Remove inside parenthesis ", replace_cb(r"[\uff08\u0028].+?[\uff09\u0029]", "") ),
            (" Remove quote mark ", replace_cb(r"[\u0022\u0027]", "") ),
            (" to single-line ", replace_cb(r"\r?\n", "") ),
            (" to single-line without whitespace ", replace_cb(r"\s", "") ),
        ],
        "Transform Alphabet": [
            (" to FullWidth ", format_cb(CharWidth().to_full_width) ),
            (" to HalfWidth ", format_cb(CharWidth().to_half_width) ),
            (" to lowercase ", lambda : get_current_clipboard().lower() ),
            (" to UPPERCASE ", lambda : get_current_clipboard().upper() ),
            (" to Title Case ", format_cb(to_smart_title_case) ),
        ],
        "Transform Single Punctuation": [
            (" Colon to double horizontal-bar ", replace_cb(r"\s*[\uff1a\u003a]\s*", "\u2015\u2015") ),
            (" Fix Word Bullet ", replace_cb(r"\uf09f\u0009", "\u30fb") ),
            (" to Official-comma ", replace_cb(r"\u3001", "\uff0c") ),
            (" to Traditional-comma ", replace_cb(r"\uff0c", "\u3001") ),
        ],
        "Transform Paired Punctuation": [
            (" Parenthesis: to FullWidth ", format_cb(CharWidth(True).to_full_width) ),
            (" Parenthesis: to HalfWidth ", format_cb(CharWidth(True).to_half_width) ),
            (" Parenthesis: to Tortoise ", format_cb(to_tortoise) ),
            (" Quotation: fix Dumb ", format_cb(fix_dumb_quotation) ),
            (" Quotation: Escape ", replace_cb(r'"', r'\"')),
            (" Quotation: to Double ", format_cb(to_double)),
            (" Quotation: to Single ", format_cb(to_single) ),
        ],
        "URL Format": [
            (" Decode URL ", format_cb(decode_url) ),
            (" Embed Youtube-Movie (Native) ", replace_cb(r"^.+\.youtube\.com/watch\?v=(.{11}).*", r'<div class="youtube"><iframe loading="lazy" src="https://www.youtube.com/embed/\1?mute=1&rel=0" frameborder="0" allowfullscreen></iframe></div>') ),
            (" Embed Youtube-Movie (Lite) ", replace_cb(r"^.+\.youtube\.com/watch\?v=(.{11}).*", r'<lite-youtube videoid="\1" playlabel="" params="controls=1&rel=0&mute=1"></lite-youtube>') ),
            (" Shorten Amazon-URL ", replace_cb(r"^.+amazon\.co\.jp/.+dp/(.{10}).*", r"https://www.amazon.jp/dp/\1") ),
        ],
        "Others": [
            (" Insert blank-line ", replace_cb(r"(\r?\n)+", os.linesep * 2) ),
            (" Cat local file ", format_cb(catanate_file_content) ),
            (" Fix KANGXI RADICALS ", format_cb(format_kangxi_radicals) ),
            (" Trim postalcode from address ", format_cb(strip_postalcode_from_address) ),
            (" Zoom invitation ", format_cb(format_zoom_invitation) ),
        ]
    }.items():
        m = menu + [("--------------- EXIT ---------------", lambda : None)]
        keymap.cblisters += [(title, cblister_FixedPhrase(m))]


