from . import bind_core  # noqa: N999


def configure(keymap) -> None:

    # user modifier
    keymap.replace_key("(29)", 235)  # "muhenkan" => 235
    keymap.replace_key("(28)", 236)  # "henkan" => 236
    keymap.define_modifier(235, "User0")  # "muhenkan" => "U0"
    keymap.define_modifier(236, "User1")  # "henkan" => "U1"

    # --- clipboard history ---------------------------------------------
    keymap.clipboard_history.max_items = 500
    keymap.clipboard_history.max_data_size = 10 * 1024 * 1024

    bind_core.bind(keymap)
