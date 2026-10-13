import clipboard  # ty: ignore[unresolved-import]
import core  # ty: ignore[unresolved-import]
import ime  # ty: ignore[unresolved-import]
import mouse_cursor  # ty: ignore[unresolved-import]
import sender  # ty: ignore[unresolved-import]
import web_search  # ty: ignore[unresolved-import]
import wnd_activate  # ty: ignore[unresolved-import]
import wnd_snap  # ty: ignore[unresolved-import]
from keytables import (  # ty: ignore[unresolved-import]
    browser,
    excel,
    mery,
    slack,
    smoothcsv,
    sumatra_pdf,
    vscode,
)


def configure(keymap) -> None:
    """
    https://github.com/crftwr/keyhac/blob/main/keyhac/_config.py
    """

    # user modifier
    keymap.replace_key("(29)", 235)  # "muhenkan" => 235
    keymap.replace_key("(28)", 236)  # "henkan" => 236
    keymap.define_modifier(235, "User0")  # "muhenkan" => "U0"
    keymap.define_modifier(236, "User1")  # "henkan" => "U1"

    # clipboard history
    keymap.clipboard_history.max_items = 500
    keymap.clipboard_history.max_data_size = 10 * 1024 * 1024

    # key bingings for every window
    kt_global = keymap.define_keytable(focus_path_pattern="*")
    for module in [
        core,
        clipboard,
        ime,
        sender,
        mouse_cursor,
        web_search,
        wnd_activate,
        wnd_snap,
    ]:
        module.bind(keymap, kt_global)

    # key bingings for specific app window
    for kt in [
        browser,
        mery,
        slack,
        smoothcsv,
        sumatra_pdf,
        vscode,
        excel,
    ]:
        kt.bind(keymap)
