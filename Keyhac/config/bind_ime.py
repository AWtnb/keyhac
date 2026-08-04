from keyhac import *  # type: ignore

from .tools import ime as ime_tool
from .tools.common import is_global_target


def bind(keymap) -> None:
    ime_tool.setup(keymap)
    km = keymap.defineWindowKeymap(check_func=is_global_target)

    ime_handler = ime_tool.Handler(0)
    for key, func in {
        "U1-J": ime_handler.to_skk_kana,
        "LC-U0-I": ime_handler.to_skk_kata,
        "U0-F7": ime_handler.to_skk_kata,
        "U0-O": ime_handler.to_skk_half_kata,
        "LC-LS-U0-I": ime_handler.to_skk_half_kata,
        "U0-F8": ime_handler.to_skk_half_kata,
        "U0-F": ime_tool.disable,
        "LS-U0-F": ime_handler.to_skk_kana,
        "S-U1-J": ime_handler.to_skk_latin,
        "U1-I": ime_handler.reconvert_with_skk,
        "O-(236)": ime_handler.to_skk_abbrev,
        "U1-U": ime_handler.start_skk_conv_suffix,
    }.items():
        km[key] = func
