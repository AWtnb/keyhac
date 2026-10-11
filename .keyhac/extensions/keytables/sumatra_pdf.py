from libs import sender


def bind(keymap) -> None:
    sender.setup(keymap)
    kt = keymap.define_keytable(
        custom_condition_func=(
            lambda focus: focus.app_name == "SumatraPDF" and focus.class_name != "Edit"
        )
    )

    for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        kt[key] = sender.DirectSender().invoke(key)
