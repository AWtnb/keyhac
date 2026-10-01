import shutil
import subprocess
import webbrowser
from collections.abc import Callable

from keyhac import ThreadedAction  # ty: ignore[unresolved-import]
from libs import clipboard
from libs.text_utils.commands import CLIPBOARD_FORMAT_COMMANDS
from libs.text_utils.misc import (
    as_single_quoted_line,
    make_str_cleaner,
    simple_quote,
    to_full_letter,
    to_half_letter,
)


def bind(keymap) -> None:

    clipboard.setup(keymap)

    kt = keymap.define_keytable(focus_path_pattern="*")

    kt["U0-V"] = clipboard.Paste()

    def bind_cleanup_paster(key: str):
        kt_paster = keymap.define_keytable(name="paste with whitespace trimmed:")
        for mod1, no_space in {
            "": False,
            "C-": True,
        }.items():
            for mod2, as_single_line in {
                "": False,
                "S-": True,
            }.items():
                cleaner = make_str_cleaner(no_space, as_single_line)
                kt_paster[mod1 + mod2 + key] = clipboard.Paste(format_func=cleaner)
        return kt_paster

    kt["U1-V"] = bind_cleanup_paster("V")

    # paste with quote mark
    kt["U1-Q"] = clipboard.Paste(format_func=simple_quote)
    kt["LC-U1-Q"] = clipboard.Paste(format_func=as_single_quoted_line)

    # paste as fullwidth / halfwidth
    kt["U1-W"] = clipboard.Paste(format_func=lambda s: to_full_letter(s, True))
    kt["LS-U1-W"] = clipboard.Paste(format_func=lambda s: to_half_letter(s, True))

    def open_selected_url(url: str) -> None:
        try:
            webbrowser.open(url.strip())
        except Exception as e:  # noqa: BLE001
            print(e)

    kt["C-U0-O"] = clipboard.CopyThen(open_selected_url)

    class FuzzyClipboadCommands(ThreadedAction):
        def run(self) -> Callable | None:
            if not shutil.which("fzf"):
                return None

            if not clipboard.get_string():
                return None

            proc = subprocess.Popen(
                ["fzf.exe", "--no-mouse", "--margin=1"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                encoding="utf-8",
            )
            try:
                if proc.stdin:
                    for k in CLIPBOARD_FORMAT_COMMANDS:
                        proc.stdin.write(k + "\n")
                    proc.stdin.close()
            except Exception:  # noqa: BLE001
                print("Failed to write to fzf stdin")
                return None

            result, err = proc.communicate()
            if proc.returncode != 0:
                if err:
                    print(err)
                return None
            result = result.strip()
            if len(result) < 1:
                return None

            return CLIPBOARD_FORMAT_COMMANDS.get(result)

        def finished(self, selected: Callable | None) -> None:
            if selected is not None:
                clipboard.Paste(None, selected)()

    kt["U1-Z"] = FuzzyClipboadCommands()
