from keyhac import SnapWindow  # ty: ignore[unresolved-import]


def bind(keymap) -> None:
    kt = keymap.define_keytable(focus_path_pattern="*")
    kt["U1-H"] = "LWin-Left"
    kt["U1-L"] = "LWin-Right"

    kt_snap = keymap.define_keytable()
    kt_snap["U0-L"] = "LWin-Shift-Right"
    kt_snap["U0-H"] = "LWin-Shift-Left"

    def minimize() -> None:
        wnd = keymap.get_active_window()
        if wnd:
            wnd.minimize()

    kt_snap["N"] = minimize

    kt_snap["X"] = SnapWindow(position="full")  # TODO: maximize

    for mod, ratio in {
        "": 1 / 2,
        "S-": 2 / 3,
        "C-": 1 / 3,
    }.items():
        for key, position in {
            "H": "left",
            "L": "right",
            "J": "bottom",
            "K": "top",
        }.items():
            kt_snap[mod + key] = SnapWindow(position=position, ratio=ratio)

    class HalfSizer:
        def __init__(self, direction: str) -> None:
            if direction not in ("left", "right", "up", "down"):
                raise ValueError(f"invalid direction: {direction}")

            self.direction = direction

        def __call__(self) -> None:
            active_wnd = keymap.get_active_window()
            if active_wnd is None:
                return
            print(active_wnd.get_frame())

    for key, direction in {
        "H": "left",
        "L": "right",
        "J": "down",
        "K": "up",
    }.items():
        kt_snap[f"U1-{key}"] = HalfSizer(direction)

    kt["U1-M"] = kt_snap
