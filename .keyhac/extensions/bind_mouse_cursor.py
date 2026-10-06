from libs import mouse_cursor


def bind(keymap) -> None:
    mouse_cursor.setup(keymap)

    kt = keymap.define_keytable(focus_path_pattern="*")

    def fetch_cursor() -> None:
        print(keymap.get_active_window().get_frame())
        x, y, w, h = keymap.get_active_window().get_frame()
        center_x = x + w // 2
        center_y = y + h // 2
        mouse_cursor.set_position(center_x, center_y)

    kt["O-RShift"] = fetch_cursor
