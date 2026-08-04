import os
import subprocess
import webbrowser

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
from .tools.clipboard import copy_then, invoke_clean_paster, paste
from .tools.common import (
    CallbackFunc,
    balloon,
    check_fzf,
    delay,
    get_now,
    is_global_target,
    is_keyhac_console,
    open_vscode,
    shell_exec,
    smart_check_path,
)
from .tools.format_clipboard import CLIPBOARD_FORMATTER_MAPPING
from .tools.format_str import (
    as_single_quoted_line,
    simple_quote,
    to_full_letter,
    to_half_letter,
)
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

    def bind_ime_handler() -> None:
        ime_handler = ime_tool.Handler(0)
        for key, func in {
            "U1-J": ime_handler.to_skk_kana,
            "LC-U0-I": ime_handler.to_skk_kata,
            "U0-F7": ime_handler.to_skk_kata,
            "U0-O": ime_handler.to_skk_half_kata,
            "LC-LS-U0-I": ime_handler.to_skk_half_kata,
            "U0-F8": ime_handler.to_skk_half_kata,
            "U0-F": ime_tool.disable,
            "LS-U0-F": ime_handler.to_skk_kana,
            "S-U1-J": ime_handler.to_skk_latin,
            "U1-I": ime_handler.reconvert_with_skk,
            "O-(236)": ime_handler.to_skk_abbrev,
            "U1-U": ime_handler.start_skk_conv_suffix,
        }.items():
            keymap_global[key] = func

    bind_ime_handler()

    keymap_global["U0-V"] = paste

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
    keymap_global["U1-Q"] = lambda: paste(None, simple_quote)
    keymap_global["LC-U1-Q"] = lambda: paste(None, as_single_quoted_line)

    # paste as fullwidth / halfwidth
    keymap_global["U1-W"] = lambda: paste(format_func=lambda s: to_full_letter(s, True))
    keymap_global["LS-U1-W"] = lambda: paste(
        format_func=lambda s: to_half_letter(s, True)
    )

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

        copy_then(_open)

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
                sender.ime_handler.to_skk_full_latin, symbol, ime_tool.SKKKey.kana
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
                sender.ime_handler.to_skk_full_latin,
                *pair,
                "Left",
                ime_tool.SKKKey.kana,
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

            copy_then(_search)

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

    SYSTEM_BROWSER = SystemBrowser()

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

    keymap_global["U0-Q"] = search_on_browser

    ################################
    # popup clipboard menu
    ################################

    def fzfmenu() -> None:
        if not check_fzf():
            balloon(keymap, "cannot find fzf on PC.")
            return

        if not cb_tool.get_string():
            balloon(keymap, "no text in clipboard.")
            return

        table = CLIPBOARD_FORMATTER_MAPPING

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
                    for k in table:
                        proc.stdin.write(k + "\n")
                    proc.stdin.close()
            except Exception as e:  # noqa: BLE001
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
                paste(None, job_item.func)

        subthread_tool.run(_fzf, _finished, True)

    keymap_global["U1-Z"] = fzfmenu
