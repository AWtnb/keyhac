from enum import Enum, StrEnum


def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


class ImeStatus(Enum):
    on = True
    off = False


def get_status() -> ImeStatus | None:
    return keymap.get_ime_status()


def set_status(status: ImeStatus) -> bool:
    return keymap.set_ime_status(status.value)


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
    IMEを有効にしてキー入力でSKKKを初期化する
    """
    with keymap.get_input_context() as ctx:
        ctx.send_key(SKKKey.kana)
        for key in seq:
            ctx.send_key(key)


def turnon_skk() -> bool:
    status = get_status()
    if status is None:
        return False
    if status == ImeStatus.on:
        return True
    return set_status(ImeStatus.on)


def turnoff_skk() -> None:
    status = get_status()
    if status:
        _send_ime_sequence(SKKKey.toggle_vk)


def to_skk_kana() -> None:
    if turnon_skk():
        _send_ime_sequence()


def to_skk_latin() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.latin)


def to_skk_abbrev() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.abbrev)


def to_skk_kata() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.kata)


def to_skk_half_kata() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.halfkata)


def to_skk_full_latin() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.jlatin)


def start_skk_conv() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.convpoint)


def start_skk_conv_suffix() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.convpoint, SKKKey.affix)


def reconvert_with_skk() -> None:
    if turnon_skk():
        _send_ime_sequence(SKKKey.reconv, SKKKey.cancel)
