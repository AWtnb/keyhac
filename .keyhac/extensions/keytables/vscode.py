from libs import ime, sender


def bind(keymap) -> None:
    sender.setup(keymap)

    kt = keymap.define_keytable(app="Code.exe")
    kt["U0-Slash"] = "C-Slash", "A-S-Down", "C-Slash"

    for key in [
        "C-E",
        "C-F",
        "C-T",
        "C-S-F",
        "C-S-E",
        "C-S-O",
        "C-S-G",
        "RC-RS-X",
        "C-0",
        "C-S-P",
        "C-A-B",
        "C-A-AtMark",
        "C-1",
        "C-2",
        "C-S-Enter",
        "S-Enter",
    ]:
        kt[key] = sender.SendThen(ime.turnoff_skk, [key])
