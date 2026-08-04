from enum import Enum, StrEnum

from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import virtual_finger
from .virtual_finger import as_motion_sequence


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)


class ImeStatus(Enum):
    on = 1
    off = 0


def get_status() -> ImeStatus:
    return ImeStatus(keymap.getWindow().getImeStatus())


def set_status(status: ImeStatus) -> None:
    keymap.getWindow().setImeStatus(status.value)


def is_enabled() -> bool:
    return get_status() == ImeStatus.on


def enable() -> None:
    if not is_enabled():
        set_status(ImeStatus.on)


# Unlike the `turnoff_skk` method, this method forcibly turns off the IME itself.
# Once SKK is disabled with this method, the next execution of the `enable` method starts SKK with the mode it was in just before being turned off.
def disable() -> None:
    if is_enabled():
        set_status(ImeStatus.off)


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


TO_KANA_SEQ = as_motion_sequence(SKKKey.kana)
TO_TURNOFF_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.toggle_vk)
TO_KATA_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.kata)
TO_LATIN_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.latin)
TO_ABBREV_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.abbrev)
TO_HALF_KATA_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.halfkata)
TO_FULL_LATIN_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.jlatin)
TO_CONV_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.convpoint)
TO_CONV_SUFFIX_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.convpoint, SKKKey.affix)
TO_RECONV_SEQ = as_motion_sequence(SKKKey.kana, SKKKey.reconv, SKKKey.cancel)


class Handler:
    def __init__(self, inter_stroke_pause: int = 10) -> None:
        self._finger = virtual_finger.VirtualFinger(inter_stroke_pause)

    def turnoff_skk(self) -> None:
        if is_enabled():
            self._finger.send_motion_sequence(*TO_TURNOFF_SEQ)

    def to_skk_kana(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_KANA_SEQ)

    def to_skk_latin(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_LATIN_SEQ)

    def to_skk_abbrev(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_ABBREV_SEQ)

    def to_skk_kata(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_KATA_SEQ)

    def to_skk_half_kata(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_HALF_KATA_SEQ)

    def to_skk_full_latin(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_FULL_LATIN_SEQ)

    def start_skk_conv(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_CONV_SEQ)

    def start_skk_conv_suffix(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_CONV_SUFFIX_SEQ)

    def reconvert_with_skk(self) -> None:
        enable()
        self._finger.send_motion_sequence(*TO_RECONV_SEQ)
