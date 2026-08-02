from enum import Enum

from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore  # noqa: F403

from . import virtual_finger
from .virtual_finger import Tap

keymap: WindowKeymap = None
virtual_finger.keymap = keymap


class SKKKey:
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


class ImeStatus(Enum):
    on = 1
    off = 0


class ImeControl:
    def __init__(self, inter_stroke_pause: int = 10) -> None:
        self._finger = virtual_finger.VirtualFinger(inter_stroke_pause)

        self.taps_to_kana = self._tapify()
        self.taps_to_turnoff = self._tapify(SKKKey.toggle_vk)
        self.taps_to_kata = self._tapify(SKKKey.kata)
        self.taps_to_latin = self._tapify(SKKKey.latin)
        self.taps_to_abbrev = self._tapify(SKKKey.abbrev)
        self.taps_to_half_kata = self._tapify(SKKKey.halfkata)
        self.taps_to_full_latin = self._tapify(SKKKey.jlatin)
        self.taps_to_conv = self._tapify(SKKKey.convpoint)
        self.taps_to_conv_suffix = self._tapify(SKKKey.convpoint, SKKKey.affix)
        self.taps_to_reconv = self._tapify(SKKKey.reconv, SKKKey.cancel)

    def _tapify(self, *keys: str) -> list[Tap]:
        return self._finger.compile(SKKKey.kana, *keys)

    @staticmethod
    def get_status() -> ImeStatus:
        return ImeStatus(keymap.getWindow().getImeStatus())

    @staticmethod
    def set_status(status: ImeStatus) -> None:
        keymap.getWindow().setImeStatus(status.value)

    @classmethod
    def is_enabled(cls) -> bool:
        return cls.get_status() == ImeStatus.on

    @classmethod
    def enable(cls) -> None:
        if not cls.is_enabled():
            cls.set_status(ImeStatus.on)

    # Unlike the `turnoff_skk` method, this method forcibly turns off the IME itself.
    # Once SKK is disabled with this method, the next execution of the `enable` method starts SKK with the mode it was in just before being turned off.
    @classmethod
    def disable(cls) -> None:
        if cls.is_enabled():
            cls.set_status(ImeStatus.off)

    def turnoff_skk(self) -> None:
        if self.is_enabled():
            self._finger.send_compiled(*self.taps_to_turnoff)

    def to_skk_kana(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_kana)

    def to_skk_latin(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_latin)

    def to_skk_abbrev(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_abbrev)

    def to_skk_kata(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_kata)

    def to_skk_half_kata(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_half_kata)

    def to_skk_full_latin(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_full_latin)

    def start_skk_conv(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_conv)

    def start_skk_conv_suffix(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_conv_suffix)

    def reconvert_with_skk(self) -> None:
        self.enable()
        self._finger.send_compiled(*self.taps_to_reconv)
