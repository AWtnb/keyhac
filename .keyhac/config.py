import bind_app_specific  # ty: ignore[unresolved-import]
import bind_clipboard  # ty: ignore[unresolved-import]
import bind_core  # ty: ignore[unresolved-import]
import bind_ime  # ty: ignore[unresolved-import]
import bind_input  # ty: ignore[unresolved-import]
import bind_mouse_cursor  # ty: ignore[unresolved-import]
import bind_wnd_activate  # ty: ignore[unresolved-import]


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

    # key bingings
    bind_core.bind(keymap)
    bind_ime.bind(keymap)
    bind_clipboard.bind(keymap)
    bind_app_specific.bind(keymap)
    bind_wnd_activate.bind(keymap)
    bind_input.bind(keymap)
    bind_mouse_cursor.bind(keymap)
