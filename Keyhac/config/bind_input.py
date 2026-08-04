from keyhac import *  # type: ignore

from .tools import ime as ime_tool
from .tools import sender as sender_tool
from .tools.common import is_global_target


def bind(keymap) -> None:
    ime_tool.setup(keymap)
    sender_tool.setup(keymap)

    km = keymap.defineWindowKeymap(check_func=is_global_target)

    base_sender = sender_tool.SKKSender()

    km["Yen"] = base_sender.invoke_emitThen(ime_tool.ImeStatus.off, "Yen")
    km["U0-P"] = base_sender.under_kanamode("・")
    km["U0-AtMark"] = base_sender.invoke_emitThen(
        ime_tool.ImeStatus.off, "LS-AtMark", "LS-AtMark", "Left"
    )
    km["U0-5"] = base_sender.invoke_emitThen(
        ime_tool.ImeStatus.off, "S-7", "S-5", "S-5", "S-7", "Left", "Left"
    )

    # select-to-left with ime control
    km["U1-B"] = base_sender.under_kanamode("S-Left")
    km["LS-U1-B"] = base_sender.under_kanamode("S-Right")
    km["U1-Space"] = base_sender.under_kanamode("C-S-Left")
    km["U1-4"] = base_sender.under_kanamode(ime_tool.SKKKey.convpoint, "S-4", "Tab")

    for key, symbol in {
        "S-U0-Colon": "\uff1a",  # FULLWIDTH COLON
        "S-U0-Comma": "\uff0c",  # FULLWIDTH COMMA
        "S-U0-Period": "\uff0e",  # FULLWIDTH PERIOD
    }.items():
        km[key] = base_sender.invoke(
            base_sender.ime_handler.to_skk_full_latin, symbol, ime_tool.SKKKey.kana
        )

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
        km[key] = base_sender.invoke(
            base_sender.ime_handler.to_skk_full_latin,
            *pair,
            "Left",
            ime_tool.SKKKey.kana,
        )

    direct_sender = sender_tool.DirectSender()
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
        "S-U0-8": ("U-Shift", "Minus", "Space", ime_tool.SKKKey.toggle_vk),
        "U1-1": ("1.", "Space", ime_tool.SKKKey.toggle_vk),
        "S-U0-SemiColon": ("U-Shift", "SemiColon"),
        "U0-T": ("</>", "Left", "S-Left"),
    }.items():
        km[key] = direct_sender.invoke(*sent)

    for key, circumfix in {
        "U0-CloseBracket": ["[", "]"],
        "U1-9": ["(", ")"],
        "S-U0-9": ['("', '")'],
        "U1-CloseBracket": ["{", "}"],
    }.items():
        _, suffix = circumfix
        sequence = circumfix + ["Left"] * len(suffix)
        km[key] = direct_sender.invoke(*sequence)
