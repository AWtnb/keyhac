import os
import subprocess
import webbrowser

import ckit  # type: ignore
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
from .tools.clipboard import copy_then, invoke_clean_paster, paste
from .tools.common import (
    balloon,
    check_fzf,
    delay,
    get_now,
    is_global_target,
    open_vscode,
    shell_exec,
    smart_check_path,
)
from .tools.format_clipboard import CLIPBOARD_FORMATTER_MAPPING
from .tools.str_tools import (
    as_single_quoted_line,
    simple_quote,
    to_full_letter,
    to_half_letter,
)


def setup(keymap) -> None:

    vf_tool.setup(keymap)
    subthread_tool.setup(keymap)
    ime_tool.setup(keymap)
    sender_tool.setup(keymap)
    cb_tool.setup(keymap)
    cursor_pos_tool.setup(keymap)
    window_snap_tool.setup(keymap)
    window_activate_tool.setup(keymap)

    # keymap working on any window
    keymap_global = keymap.defineWindowKeymap(check_func=is_global_target)

    def safe_close() -> None:
        finger = vf_tool.VirtualFinger(0)
        close_key = finger.compile("A-F4")

        def _wait(_) -> None:
            delay(200)

        def _close(_) -> None:
            finger.send_compiled(*close_key)

        subthread_tool.run(_wait, _close)

    keymap_global["C-Q"] = safe_close

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
