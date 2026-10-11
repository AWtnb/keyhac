from libs import sender


def bind(keymap) -> None:
    sender.setup(keymap)
    kt = keymap.define_keytable(app="excel.exe")

    def select_all() -> None:
        focus = keymap.focus
        if focus is None:
            return

        sequence = ["C-End", "C-S-Home"] if focus.class_name == "EXCEL6" else ["C-A"]
        sender.send_keys(*sequence)

    kt["C-A"] = select_all
