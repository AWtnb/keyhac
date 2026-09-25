from collections.abc import Callable

from . import ime as ime_tool
from ._common import CallbackFunc, delay


def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap
    ime_tool.setup(keymap)


def send_keys(*keys: str) -> None:
    with keymap.get_input_context() as ctx:
        for key in keys:
            ctx.send_key(key)


def _send_sequence(inter_stroke_pause: int, *sequence: str) -> None:
    with keymap.get_input_context() as ctx:
        for key in sequence:
            delay(inter_stroke_pause)
            try:
                ctx.send_key(key)
            except (ValueError, KeyError):
                ctx.send_text(key)


class SKKSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self._inter_stroke_pause = inter_stroke_pause

    def invoke(self, mode_setter: Callable[[], bool], *sequence: str) -> CallbackFunc:

        def _sender() -> None:
            if mode_setter():
                _send_sequence(self._inter_stroke_pause, *sequence)

        return _sender

    def under_kanamode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(ime_tool.to_skk_kana, *sequence)
        return sender

    def under_latinmode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(ime_tool.to_skk_latin, *sequence)
        return sender

    def without_mode(self, *sequence: str) -> CallbackFunc:
        sender = self.invoke(ime_tool.turnoff_skk, *sequence)
        return sender

    def invoke_emitThen(self, later_ime_status: bool, *sequence: str) -> CallbackFunc:

        def _sender() -> None:
            status = ime_tool.get_status()
            if status is None:
                return
            seq = list(sequence)
            if status != later_ime_status:
                seq.append(ime_tool.SKKKey.toggle_vk)
            _send_sequence(self._inter_stroke_pause, *seq)

        return _sender


class DirectSender:
    def __init__(self, inter_stroke_pause: int = 0) -> None:
        self.skk = SKKSender(inter_stroke_pause)

    def invoke(self, *sequence: str) -> CallbackFunc:
        seq = list(sequence)
        return self.skk.invoke(ime_tool.turnoff_skk, *seq)
