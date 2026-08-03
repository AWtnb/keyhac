import os
import re
import subprocess
import unicodedata
import urllib.parse
import webbrowser
from collections.abc import Callable

import ckit  # type: ignore
import pyauto  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from .tools import clipboard as cb_tool
from .tools import cursor_pos as cursor_pos_tool
from .tools import ime as ime_tool
from .tools import sender as sender_tool
from .tools import subthread as subthread_tool
from .tools import virtual_finger as vf_tool
from .tools import window_activate as window_activate_tool
from .tools import window_snap as window_snap_tool
from .tools.browser_info import SystemBrowser
from .tools.clipboard import (
    invoke_clean_paster,
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
from .tools.punctuation import KANGXI_RADICAL_MAPPING, RADICAL_MAPPING
from .tools.web_search import invoke_web_seacher
from .tools.window_activate import WindowActivator, WndScanner
from .tools.window_rect import RectEdge
from .tools.window_snap import invoke_maximized_snapper, invoke_shrinker, invoke_snapper


def setup(keymap) -> None:

    vf_tool.setup(keymap)
    subthread_tool.keymap = keymap
    ime_tool.keymap = keymap
    sender_tool.keymap = keymap
    cb_tool.setup(keymap)
    cursor_pos_tool.setup(keymap)
    window_snap_tool.setup(keymap)
    window_activate_tool.setup(keymap)

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

    keymap_global["U0-V"] = cb_tool.Manager.paste

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
    keymap_global["U1-Q"] = lambda: cb_tool.Manager().paste(None, simple_quote)
    keymap_global["LC-U1-Q"] = lambda: cb_tool.Manager().paste(None, as_single_line)
    keymap_global["LS-U1-Q"] = lambda: cb_tool.Manager().paste(None, skip_blank_line)

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

    def bind_window_snapper(km: WindowKeymap) -> None:
        altkey_stat = {0: "", 1: "LA-", 2: "RA-"}
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

        for idx, alt in altkey_stat.items():
            for area_mod, scale in scale_mapping.items():
                for key, edge in edge_mapping.items():
                    km[alt + area_mod + key] = invoke_snapper(idx, scale, edge)

    bind_window_snapper(keymap_global["U1-M"])

    def bind_maximized_window_snapper() -> None:
        for key in ["0", "1", "2"]:
            monitor_idx = int(key)
            _snap = invoke_maximized_snapper(monitor_idx)
            keymap_global["U1-M"][str(key)] = _snap

    bind_maximized_window_snapper()

    def bind_window_shrinker(km: WindowKeymap) -> None:
        for key, toward in {
            "H": RectEdge.left,
            "L": RectEdge.right,
            "K": RectEdge.top,
            "J": RectEdge.bottom,
        }.items():
            km["U1-" + key] = invoke_shrinker(toward)

    bind_window_shrinker(keymap_global["U1-M"])

    ################################
    # set cursor position
    ################################

    keymap_global["O-RCtrl"] = cursor_pos_tool.snap_cursor
    keymap_global["O-RShift"] = cursor_pos_tool.snap_to_center

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

    def invoke_web_search_job(
        uri: str, strict: bool = False, strip_hiragana: bool = False
    ) -> CallbackFunc:
        search_func = invoke_web_seacher(uri, strict, strip_hiragana)

        def _searcher() -> None:
            def _search(job_item: ckit.JobItem) -> None:
                s = job_item.copied
                if len(s) < 1:
                    s = job_item.origin
                search_func(s)

            cb_tool.Manager().after_register(_search)

        return _searcher

    def bind_web_search_key(km: WindowKeymap, mapping: dict[str, str]) -> None:
        for shift_key in ("", "S-"):
            for ctrl_key in ("", "C-"):
                is_strict = shift_key != ""
                strip_hiragana = ctrl_key != ""
                trigger_key = shift_key + ctrl_key + "U0-S"
                km[trigger_key] = keymap.defineMultiStrokeKeymap()
                for key, uri in mapping.items():
                    km[trigger_key][key] = invoke_web_search_job(
                        uri, is_strict, strip_hiragana
                    )

    bind_web_search_key(
        keymap_global,
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
        },
    )

    ################################
    # activate window
    ################################

    keymap.default_browser = SystemBrowser()

    def bind_cute_exec(wnd_keymap: WindowKeymap, remap_table: dict) -> None:
        for key, params in remap_table.items():
            func = window_activate_tool.invoke_cute_exec(*params)
            wnd_keymap[key] = func

    bind_cute_exec(
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
    bind_cute_exec(
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
            "fix KANGXI RADICALS": lambda s: s.transrate(
                str.maketrans(KANGXI_RADICAL_MAPPING | RADICAL_MAPPING)
            ),
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
