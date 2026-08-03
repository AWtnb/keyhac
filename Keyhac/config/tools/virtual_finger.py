import pyauto  # type: ignore
from keyhac_keymap import KeyCondition, WindowKeymap  # ty:ignore[unresolved-import]

from .common import delay


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


class Tap:
    mod: int = 0
    sequence: list[pyauto.Key | pyauto.KeyUp | pyauto.KeyDown | pyauto.Char] = []  # noqa: RUF012

    def __init__(self, name: str):
        up = None
        tokens = [s for s in name.split("-")]

        for token in tokens[:-1]:
            t = token.strip().upper()
            try:
                self.mod |= KeyCondition.strToMod(t, force_LR=True)
            except ValueError:
                if up is not None:
                    continue
                if t == "U":
                    up = True
                else:
                    if t == "D":
                        up = False

        tail = tokens[-1]
        try:
            vk = KeyCondition.strToVk(tail.strip().upper())
            if up is None:
                self.sequence = [pyauto.Key(vk)]
            else:
                if up:
                    self.sequence = [pyauto.KeyUp(vk)]
                else:
                    self.sequence = [pyauto.KeyDown(vk)]
        except ValueError:
            self.sequence = [pyauto.Char(c) for c in str(tail)]


class VirtualFinger:
    def __init__(self, inter_stroke_pause: int = 10) -> None:

        self._inter_stroke_pause = inter_stroke_pause

        mod_keys = ["Shift", "Alt", "Ctrl"]
        self._mod_release_taps = [Tap(f"U-{mod}") for mod in mod_keys]

    @staticmethod
    def begin() -> None:
        keymap.beginInput()
        keymap.setInput_Modifier(0)

    @staticmethod
    def end() -> None:
        keymap.endInput()

    @staticmethod
    def compile(*sequence: str) -> list[Tap]:
        return [Tap(elem) for elem in sequence]

    def send(self, *sequence: str) -> None:
        taps = self.compile(*sequence)
        self.send_compiled(*taps)

    def send_compiled(self, *taps: Tap) -> None:
        self.begin()

        for t in self._mod_release_taps:
            for x in t.sequence:
                keymap.input_seq.append(x)

        for t in taps:
            delay(self._inter_stroke_pause)
            keymap.setInput_Modifier(t.mod)
            for x in t.sequence:
                keymap.input_seq.append(x)

        self.end()
