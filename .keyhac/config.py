import bind_app_specific  # ty: ignore[unresolved-import]
import bind_clipboard  # ty: ignore[unresolved-import]
import bind_core  # ty: ignore[unresolved-import]
import bind_ime  # ty: ignore[unresolved-import]
import bind_input  # ty: ignore[unresolved-import]
import bind_mouse_cursor  # ty: ignore[unresolved-import]
import bind_web_search  # ty: ignore[unresolved-import]
import bind_wnd_activate  # ty: ignore[unresolved-import]
import bind_wnd_snap  # ty: ignore[unresolved-import]


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
        bind_core,
        bind_clipboard,
        bind_ime,
        bind_input,
        bind_mouse_cursor,
        bind_web_search,
        bind_wnd_activate,
        bind_wnd_snap,
    ]:
        module.bind(keymap, kt_global)

    # key bingings for specific app window
    bind_app_specific.bind(keymap)
