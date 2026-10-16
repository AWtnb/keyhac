import webbrowser

from keyhac import ShowClipboardTools  # ty: ignore[unresolved-import]
from libs import clipboard
from libs.text_utils.commands import CLIPBOARD_FORMAT_COMMANDS
from libs.text_utils.misc import (
    as_single_quoted_line,
    make_str_cleaner,
    simple_quote,
    to_full_letter,
    to_half_letter,
)


def bind(keymap, kt) -> None:

    clipboard.setup(keymap)

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

    # paste as code
    kt["U0-C-AtMark"] = clipboard.Paste(None, lambda s: f"`{s.strip()}`")

    def _as_codeblock(s: str) -> str:
        code = s.rstrip()
        lines = list(reversed(code.splitlines()))
        while lines:
            line = lines.pop()
            if line.strip():
                lines.append(line)
                break
        return f"\n\n```\n{'\n'.join(reversed(lines))}\n```\n\n"

    kt["U0-S-AtMark"] = clipboard.Paste(None, _as_codeblock)

    # paste as fullwidth / halfwidth
    kt["U1-W"] = clipboard.Paste(format_func=lambda s: to_full_letter(s, True))
    kt["LS-U1-W"] = clipboard.Paste(format_func=lambda s: to_half_letter(s, True))

    def open_selected_url(url: str) -> None:
        try:
            webbrowser.open(url.strip())
        except Exception as e:  # noqa: BLE001
            print(e)

    kt["C-U0-O"] = clipboard.CopyThen(open_selected_url)

    CLIPBOAR_TOOL_MENU = [
        ("", name, func) for name, func in CLIPBOARD_FORMAT_COMMANDS.items()
    ]
    kt["U1-Z"] = ShowClipboardTools(CLIPBOAR_TOOL_MENU)
