from typing import NamedTuple, Self

from keyhac import MouseMove  # ty: ignore[unresolved-import]


def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


class MousePos(NamedTuple):
    x: int
    y: int

    def get_delta(self, pos: Self) -> tuple[int, int]:
        delta_x = pos.x - self.x
        delta_y = pos.y - self.y
        return delta_x, delta_y


def set_position(x: int, y: int) -> bool:
    current = keymap.cursor_pos()
    if current is None:
        return False
    delta_x, delta_y = MousePos(*current).get_delta(MousePos(x, y))
    MouseMove(delta_x, delta_y)()
    return True
