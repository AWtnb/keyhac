from libs import mouse_cursor


def bind(keymap, kt) -> None:
    mouse_cursor.setup(keymap)

    def fetch_cursor() -> None:
        x, y, w, h = keymap.get_active_window().get_frame()
        center_x = x + w // 2
        center_y = y + h // 2
        mouse_cursor.set_position(center_x, center_y)

    kt["O-RShift"] = fetch_cursor

    def snap_cycle() -> None:
        current = keymap.cursor_pos()
        if current is None:
            return

        candidates = []
        frames = keymap.screen_work_frames()
        for x, y, w, h in frames:
            for n in (1, 3):
                candidates.append(
                    (
                        x + (w // 4) * n,
                        y + (h // 2),
                    )
                )

        try:
            idx = (candidates.index(current) + 1) % len(candidates)
        except ValueError:
            idx = 0

        mouse_cursor.set_position(*candidates[idx])

    kt["O-LCtrl"] = snap_cycle
