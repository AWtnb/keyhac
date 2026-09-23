from enum import StrEnum


def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


def get_status() -> bool | None:
    return keymap.get_ime_status()


def set_status(status: bool) -> bool:
    return keymap.set_ime_status(status)


class SKKKey(StrEnum):
    toggle_vk = "(243)"
    kata = "Q"
    kana = "C-J"
    halfkata = "C-O"
    latin = "S-L"
    cancel = "Esc"
    reconv = "LWin-Slash"
    abbrev = "Slash"
    convpoint = "S-0"
    jlatin = "S-Q"
    affix = "S-Period"


def _send_ime_sequence(*seq: str):
    """
    IMEを有効にしてキー入力でSKKを初期化する
    """
    with keymap.get_input_context() as ctx:
        ctx.send_key(SKKKey.kana)
        for key in seq:
            ctx.send_key(key)


def turnon_skk() -> bool:
    status = get_status()
    if status is None:
        return False
    if status:
        return True
    return set_status(True)


def turnoff_skk() -> bool:
    status = get_status()
    if status:
        _send_ime_sequence(SKKKey.toggle_vk)
    return True


def to_skk_kana() -> bool:
    if turnon_skk():
        _send_ime_sequence()
        return True
    return False


def to_skk_latin() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.latin)
        return True
    return False


def to_skk_abbrev() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.abbrev)
        return True
    return False


def to_skk_kata() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.kata)
        return True
    return False


def to_skk_half_kata() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.halfkata)
        return True
    return False


def to_skk_full_latin() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.jlatin)
        return True
    return False


def start_skk_conv() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.convpoint)
        return True
    return False


def start_skk_conv_suffix() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.convpoint, SKKKey.affix)
        return True
    return False


def reconvert_with_skk() -> bool:
    if turnon_skk():
        _send_ime_sequence(SKKKey.reconv, SKKKey.cancel)
        return True
    return False
