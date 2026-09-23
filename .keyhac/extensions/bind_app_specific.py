def bind_browser(keymap) -> None:
    kt = keymap.define_keytable(app="chrome|Google Chrome|firefox|Safari")
    kt["LC-LS-W"] = "A-Left"


def bind(keymap):
    bind_browser(keymap)
