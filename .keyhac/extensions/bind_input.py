from libs import ime, sender


def bind(keymap, kt) -> None:
    ime.setup(keymap)
    sender.setup(keymap)

    base_sender = sender.SKKSender()
    kt["U1-4"] = base_sender.under_convmode("S-4", "Tab")
    kt["U0-P"] = base_sender.under_kanamode("・")

    kt["Yen"] = sender.SendThen(ime.turnoff_skk, ["Yen"])
    kt["U0-AtMark"] = sender.SendThen(
        ime.turnoff_skk, ["LS-AtMark", "LS-AtMark", "Left"]
    )
    kt["U0-5"] = sender.SendThen(
        ime.turnoff_skk, ["S-7", "S-5", "S-5", "S-7", "Left", "Left"]
    )

    # select-to-left with ime control
    kt["U1-B"] = sender.SendThen(ime.turnon_skk, ["S-Left"])
    kt["LS-U1-B"] = sender.SendThen(ime.turnon_skk, ["S-Right"])
    kt["U1-Space"] = sender.SendThen(ime.turnon_skk, ["C-S-Left"])

    for key, symbol in {
        "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
        "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
        "S-U0-Period": "\uff0e",  # FULLWIDTH PERIOD
    }.items():
        kt[key] = base_sender.invoke(ime.to_skk_full_latin, symbol, ime.SKKKey.kana)

    for key, pair in {
        "U0-8": ["\u300e", "\u300f"],  # WHITE CORNER BRACKET 『』
        "U0-9": ["\u3010", "\u3011"],  # BLACK LENTICULAR BRACKET 【】
        "U0-OpenBracket": ["\u300c", "\u300d"],  # CORNER BRACKET 「」
        "U1-2": ["\u201c", "\u201d"],  # DOUBLE QUOTATION MARK “”
        "U1-7": ["\u2018", "\u2019"],  # SINGLE QUOTATION MARK ‘’
        "U1-8": ["\uff08", "\uff09"],  # FULLWIDTH PARENTHESIS （）
        "U0-Y": ["\u3008", "\u3009"],  # ANGLE BRACKET 〈〉
        "U1-Y": ["\u300a", "\u300b"],  # DOUBLE ANGLE BRACKET 《》
        "U1-T": ["\u3014", "\u3015"],  # TORTOISE BRACKET 〔〕
        "U1-OpenBracket": ["\uff3b", "\uff3d"],  # FULLWIDTH SQUARE BRACKET ［］
    }.items():
        kt[key] = base_sender.invoke(
            ime.to_skk_full_latin,
            *pair,
            "Left",
            ime.SKKKey.kana,
        )

    direct_sender = sender.DirectSender()
    for key, sent in {
        "Decimal": ("Period",),
        "U0-1": ("S-1",),
        "U0-Colon": ("Colon",),
        "U0-Slash": ("Slash",),
        "U1-Minus": ("Minus",),
        "LC-U0-U": ("Minus",),
        "U0-Comma": ("Comma",),
        "U0-Period": ("Period",),
        "S-U0-Enter": ("U-Shift", "Period"),
        "U0-Tab": ("Period", "BackSlash"),
        "U1-Tab": ("Period", "Period", "BackSlash"),
        "S-U0-8": ("U-Shift", "Minus", "Space", ime.SKKKey.toggle_vk),
        "U1-1": ("1.", "Space", ime.SKKKey.toggle_vk),
        "S-U0-SemiColon": ("U-Shift", "SemiColon"),
        "U0-T": ("</>", "Left", "S-Left"),
    }.items():
        kt[key] = direct_sender.invoke(*sent)

    for key, circumfix in {
        "U0-CloseBracket": ["[", "]"],
        "U1-9": ["(", ")"],
        "S-U0-9": ['("', '")'],
        "U1-CloseBracket": ["{", "}"],
    }.items():
        _, suffix = circumfix
        sequence = circumfix + ["Left"] * len(suffix)
        kt[key] = direct_sender.invoke(*sequence)
