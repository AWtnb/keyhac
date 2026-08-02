from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore  # noqa: F403

from . import ime, virtual_finger
from .common import CallbackFunc
from .virtual_finger import Tap

keymap: WindowKeymap = None

virtual_finger.keymap = keymap
ime.keymap = keymap


class SKKSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self.finger = virtual_finger.VirtualFinger(inter_stroke_pause)
        self.control = ime.ImeControl(inter_stroke_pause)

    def invoke(self, mode_setter: CallbackFunc, *sequence: str) -> CallbackFunc:
        taps = self.finger.compile(*sequence)

        def _sender() -> None:
            mode_setter()
            self.finger.send_compiled(*taps)

        return _sender

    def under_kanamode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(self.control.to_skk_kana, *sequence)
        return sender

    def under_latinmode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(self.control.to_skk_latin, *sequence)
        return sender

    def without_mode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(self.control.disable, *sequence)
        return sender

    def invoke_emitThen(
        self, later_ime_status: ime.ImeStatus, *sequence: str
    ) -> CallbackFunc:
        taps = self.finger.compile(*sequence)
        toggle_tap = Tap(ime.SKKKey.toggle_vk)

        def _sender() -> None:
            self.finger.send_compiled(*taps)
            if ime.ImeControl.get_status() != later_ime_status:
                self.finger.send_compiled(toggle_tap)

        return _sender


class DirectSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self.skk = SKKSender(inter_stroke_pause=inter_stroke_pause)

    def invoke(self, *sequence: str) -> CallbackFunc:
        seq = list(sequence)
        return self.skk.invoke(self.skk.control.turnoff_skk, *seq)

    def bind(self, km: WindowKeymap, binding: dict[str, tuple[str, ...]]) -> None:
        for key, sent in binding.items():
            km[key] = self.invoke(*sent)

    def bind_circumfix(self, km: WindowKeymap, binding: dict[str, list[str]]) -> None:
        for key, circumfix in binding.items():
            _, suffix = circumfix
            sequence = circumfix + ["Left"] * len(suffix)
            km[key] = self.invoke(*sequence)
