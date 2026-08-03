from enum import IntEnum
from typing import NamedTuple


class RectEdge(IntEnum):
    left = 0
    top = 1
    right = 2
    bottom = 3


class Rect(NamedTuple):
    left: int
    top: int
    right: int
    bottom: int
    width: int
    height: int

    def move_edge(self, toward: RectEdge, delta: int) -> list[int]:
        r = [
            self.left,
            self.top,
            self.right,
            self.bottom,
        ]
        opposite = (toward.value + 2) % 4
        r[opposite] = r[toward.value] + delta
        return r

    def resize(self, scale: float, toward: RectEdge) -> list[int]:
        if toward in [RectEdge.left, RectEdge.right]:
            dim = self.width
        else:
            dim = self.height
        delta = int(dim * scale)
        if toward in [RectEdge.right, RectEdge.bottom]:
            delta = delta * -1
        return self.move_edge(toward, delta)


def as_rect(left: int, top: int, right: int, bottom: int) -> Rect:
    return Rect(
        left=left,
        top=top,
        right=right,
        bottom=bottom,
        width=right - left,
        height=bottom - top,
    )
