from enum import Enum, StrEnum

from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import virtual_finger
from .virtual_finger import Tap


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


def as_taps(*keys: str) -> list[Tap]:
    return [Tap(k) for k in keys]


class Handler:
    def __init__(self, inter_stroke_pause: int = 10) -> None:
        self._finger = virtual_finger.VirtualFinger(inter_stroke_pause)

        self.taps_to_kana = as_taps(SKKKey.kana)
        self.taps_to_turnoff = as_taps(SKKKey.kana, SKKKey.toggle_vk)
        self.taps_to_kata = as_taps(SKKKey.kana, SKKKey.kata)
        self.taps_to_latin = as_taps(SKKKey.kana, SKKKey.latin)
        self.taps_to_abbrev = as_taps(SKKKey.kana, SKKKey.abbrev)
        self.taps_to_half_kata = as_taps(SKKKey.kana, SKKKey.halfkata)
        self.taps_to_full_latin = as_taps(SKKKey.kana, SKKKey.jlatin)
        self.taps_to_conv = as_taps(SKKKey.kana, SKKKey.convpoint)
        self.taps_to_conv_suffix = as_taps(SKKKey.kana, SKKKey.convpoint, SKKKey.affix)
        self.taps_to_reconv = as_taps(SKKKey.kana, SKKKey.reconv, SKKKey.cancel)

    def turnoff_skk(self) -> None:
        if is_enabled():
            self._finger.send_compiled(*self.taps_to_turnoff)

    def to_skk_kana(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_kana)

    def to_skk_latin(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_latin)

    def to_skk_abbrev(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_abbrev)

    def to_skk_kata(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_kata)

    def to_skk_half_kata(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_half_kata)

    def to_skk_full_latin(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_full_latin)

    def start_skk_conv(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_conv)

    def start_skk_conv_suffix(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_conv_suffix)

    def reconvert_with_skk(self) -> None:
        enable()
        self._finger.send_compiled(*self.taps_to_reconv)
