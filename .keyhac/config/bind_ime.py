from .tools import ime as ime_tool


def bind(keymap) -> None:
    ime_tool.setup(keymap)

    kt = keymap.define_keytable(focus_path_pattern="*")
    for key, func in {
        "U1-J": ime_tool.to_skk_kana,
        "U0-F7": ime_tool.to_skk_kata,
        "U0-O": ime_tool.to_skk_half_kata,
        "U0-F8": ime_tool.to_skk_half_kata,
        "U0-F": ime_tool.turnoff_skk,
        "LS-U0-F": ime_tool.to_skk_kana,
        "S-U1-J": ime_tool.to_skk_latin,
        "U1-I": ime_tool.reconvert_with_skk,
        "O-(236)": ime_tool.to_skk_abbrev,
        "U1-U": ime_tool.start_skk_conv_suffix,
    }.items():
        kt[key] = func
