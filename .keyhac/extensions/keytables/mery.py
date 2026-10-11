def bind(keymap) -> None:
    kt = keymap.define_keytable(app="Mery")

    for key, value in {
        "LA-LC-J": "LA-LC-N",
        "LA-LC-K": "LA-LC-LS-N",
        "LA-U0-J": "A-CloseBracket",
        "LA-U0-K": "A-OpenBracket",
        "LA-LC-U0-J": "A-C-CloseBracket",
        "LA-LC-U0-K": "A-C-OpenBracket",
        "LA-LS-U0-J": "A-S-CloseBracket",
        "LA-LS-U0-K": "A-S-OpenBracket",
    }.items():
        kt[key] = value
