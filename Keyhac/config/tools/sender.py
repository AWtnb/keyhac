from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import ime, virtual_finger
from .common import CallbackFunc
from .virtual_finger import as_motion_sequence


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap

    virtual_finger.setup(keymap)
    ime.setup(keymap)


class SKKSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self.finger = virtual_finger.VirtualFinger(inter_stroke_pause)
        self.ime_handler = ime.Handler(inter_stroke_pause)

    def invoke(self, mode_setter: CallbackFunc, *sequence: str) -> CallbackFunc:
        seq = as_motion_sequence(*sequence)

        def _sender() -> None:
            mode_setter()
            self.finger.send_motion_sequence(*seq)

        return _sender

    def under_kanamode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(self.ime_handler.to_skk_kana, *sequence)
        return sender

    def under_latinmode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(self.ime_handler.to_skk_latin, *sequence)
        return sender

    def without_mode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(ime.disable, *sequence)
        return sender

    def invoke_emitThen(
        self, later_ime_status: ime.ImeStatus, *sequence: str
    ) -> CallbackFunc:
        seq = as_motion_sequence(*sequence)
        toggle_seq = as_motion_sequence(ime.SKKKey.toggle_vk)

        def _sender() -> None:
            self.finger.send_motion_sequence(*seq)
            if ime.get_status() != later_ime_status:
                self.finger.send_motion_sequence(*toggle_seq)

        return _sender


class DirectSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self.skk = SKKSender(inter_stroke_pause=inter_stroke_pause)

    def invoke(self, *sequence: str) -> CallbackFunc:
        seq = list(sequence)
        return self.skk.invoke(self.skk.ime_handler.turnoff_skk, *seq)
