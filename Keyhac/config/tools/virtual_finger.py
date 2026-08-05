from typing import NamedTuple

import pyauto  # type: ignore
from keyhac_keymap import KeyCondition, WindowKeymap  # ty:ignore[unresolved-import]

from .common import delay


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


class Motion(NamedTuple):
    mod: int
    taps: list[pyauto.Key | pyauto.KeyUp | pyauto.KeyDown | pyauto.Char]


def as_motion(name: str) -> Motion:
    up = None
    tokens = [s for s in name.split("-")]

    mod = 0
    for token in tokens[:-1]:
        t = token.strip().upper()
        try:
            mod |= KeyCondition.strToMod(t, force_LR=True)
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
            return Motion(mod, [pyauto.Key(vk)])
        if up:
            return Motion(mod, [pyauto.KeyUp(vk)])
        return Motion(mod, [pyauto.KeyDown(vk)])
    except ValueError:
        return Motion(mod, [pyauto.Char(c) for c in str(tail)])


def as_motion_sequence(*sequence: str) -> list[Motion]:
    return [as_motion(s) for s in sequence]


MODKEY_RELEASE_MOTION_SEQUENCE = [as_motion(f"U-{mod}") for mod in ("Shift", "Ctrl")]
MODKEY_RELEASE_MOTION_SEQUENCE.extend(as_motion(f"{stat}-Alt") for stat in ("D", "U"))


class VirtualFinger:
    def __init__(self, inter_stroke_pause: int = 10) -> None:
        self._inter_stroke_pause = inter_stroke_pause

    def send(self, *sequence: str) -> None:
        seq = as_motion_sequence(*sequence)
        self.send_motion_sequence(*seq)

    def send_motion_sequence(self, *motion_sequence: Motion) -> None:
        keymap.beginInput()
        keymap.setInput_Modifier(0)

        for motion in MODKEY_RELEASE_MOTION_SEQUENCE:
            keymap.input_seq.extend(tap for tap in motion.taps)

        for motion in motion_sequence:
            delay(self._inter_stroke_pause)
            keymap.setInput_Modifier(motion.mod)
            keymap.input_seq.extend(tap for tap in motion.taps)

        keymap.endInput()
