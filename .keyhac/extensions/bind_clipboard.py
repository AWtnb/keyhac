import webbrowser

from libs import clipboard
from libs.str_tools import (
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

    def bind_cleanup_paster(kt, key: str) -> None:
        for mod1, no_space in {
            "": False,
            "C-": True,
        }.items():
            for mod2, as_single_line in {
                "": False,
                "S-": True,
            }.items():
                cleaner = make_str_cleaner(no_space, as_single_line)
                kt[mod1 + mod2 + key] = clipboard.Paste(format_func=cleaner)

    kt_v = keymap.define_keytable(name="U1-V")
    bind_cleanup_paster(kt_v, "V")
    kt["U1-V"] = kt_v

    # paste with quote mark
    kt["U1-Q"] = clipboard.Paste(format_func=simple_quote)
    kt["LC-U1-Q"] = clipboard.Paste(format_func=as_single_quoted_line)

    # paste as fullwidth / halfwidth
    kt["U1-W"] = clipboard.Paste(format_func=lambda s: to_full_letter(s, True))
    kt["LS-U1-W"] = clipboard.Paste(format_func=lambda s: to_half_letter(s, True))

    # kt["U1-Z"] = lambda: fzfmenu(keymap)

    def open_selected_url(url: str) -> None:
        try:
            webbrowser.open(url.strip())
        except Exception as e:  # noqa: BLE001
            print(e)

    kt["C-U0-O"] = clipboard.CopyThen(open_selected_url)
