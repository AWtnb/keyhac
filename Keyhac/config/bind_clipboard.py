import subprocess
import webbrowser

import ckit  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from .tools import clipboard as cb_tool
from .tools import subthread as subthread_tool
from .tools.clipboard import copy_then, invoke_clean_paster, paste
from .tools.common import (
    balloon,
    check_fzf,
    delay,
    is_global_target,
)
from .tools.format_clipboard import CLIPBOARD_FORMATTER_MAPPING
from .tools.str_tools import (
    as_single_quoted_line,
    simple_quote,
    to_full_letter,
    to_half_letter,
)


def fzfmenu(keymap) -> None:
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


def bind(keymap) -> None:

    subthread_tool.setup(keymap)
    cb_tool.setup(keymap)

    km = keymap.defineWindowKeymap(check_func=is_global_target)

    km["U0-V"] = paste

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

    km["U1-V"] = keymap.defineMultiStrokeKeymap()
    bind_cleanup_paster(km["U1-V"], "V")

    # paste with quote mark
    km["U1-Q"] = lambda: paste(None, simple_quote)
    km["LC-U1-Q"] = lambda: paste(None, as_single_quoted_line)

    # paste as fullwidth / halfwidth
    km["U1-W"] = lambda: paste(format_func=lambda s: to_full_letter(s, True))
    km["LS-U1-W"] = lambda: paste(format_func=lambda s: to_half_letter(s, True))

    # clipboard menu
    km["LC-LS-X"] = keymap.command_ClipboardList

    km["U1-Z"] = lambda: fzfmenu(keymap)

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

    km["C-U0-O"] = open_selected_url
