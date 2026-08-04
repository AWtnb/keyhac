from keyhac import *  # type: ignore

from .tools.common import is_global_target


def bind(keymap) -> None:
    km = keymap.defineWindowKeymap(check_func=is_global_target)

    # keyboard macro
    km["U0-0"] = keymap.command_RecordToggle
    km["S-U0-0"] = keymap.command_RecordClear
    km["U1-0"] = keymap.command_RecordPlay
    km["U1-F4"] = keymap.command_RecordPlay
    km["C-U0-0"] = keymap.command_RecordPlay

    # mouse scroll
    km["U1-Up"] = keymap.MouseWheelCommand(1.0)
    km["U1-Down"] = keymap.MouseWheelCommand(-1.0)
    km["U1-Left"] = keymap.MouseHorizontalWheelCommand(-1.0)
    km["U1-Right"] = keymap.MouseHorizontalWheelCommand(1.0)

    mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
    for mod_key in mod_keys:
        for key, value in {
            # move cursor
            "H": "Left",
            "J": "Down",
            "K": "Up",
            "L": "Right",
            # Back / Delete
            "B": "Back",
            "D": "Delete",
            # Home / End
            "A": "Home",
            "E": "End",
            # Enter
            "Space": "Enter",
        }.items():
            km[mod_key + "U0-" + key] = mod_key + value

    for key, value in {
        # focus taskbar
        "LC-U1-T": ("LWin-T"),
        # send n and space
        "LS-U0-N": ("N", "N", "Space"),
        # delete around cursor
        "U0-Back": ("Back", "Delete"),
        # delete to bol / eol
        "S-U0-B": ("S-Home", "Delete"),
        "S-U0-D": ("S-End", "Delete"),
        # escape
        "O-(235)": ("Esc"),
        "U0-X": ("Esc"),
        # line selection
        "U1-A": ("End", "S-Home"),
        # punctuation
        "U0-U": ("S-BackSlash"),
        "U1-S": ("Slash"),
        "U0-4": ("S-4", "S-BackSlash"),
        "U0-Enter": ("Period"),
        "U0-Z": ("Minus"),
        "U1-X": ("S-1"),
        # Insert line
        "U0-I": ("End", "Enter"),
        "S-U0-I": ("Home", "Enter", "Up"),
        # Context menu
        "U0-C": ("Apps"),
        "S-U0-C": ("S-Apps"),
        # rename
        "U0-N": ("F2"),
    }.items():
        km[key] = value

    for key, value in {
        "U0-2": "LS-2",
        "U0-7": "LS-7",
    }.items():
        km[key] = value, value, "Left"
