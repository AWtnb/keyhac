import pyauto  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from .common import get_monitor_infos
from .window_rect import as_rect


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


def get_pos() -> list:
    infos = get_monitor_infos()
    rects = [as_rect(*info[1]) for info in infos]
    pos = []
    for rect in rects:
        for i in (1, 3):
            y = rect.top + int(rect.height / 2)
            x = rect.left + int(rect.width / 4) * i
            pos.append([x, y])
    return pos


def set_position(x: int, y: int) -> None:
    keymap.beginInput()
    keymap.input_seq.append(pyauto.MouseMove(x, y))
    keymap.endInput()


def snap_cursor() -> None:
    pos = get_pos()
    x, y = pyauto.Input.getCursorPos()
    idx = -1
    for i, p in enumerate(pos):
        if p[0] == x and p[1] == y:
            idx = i
    if idx < 0 or idx == len(pos) - 1:
        set_position(*pos[0])
    else:
        set_position(*pos[idx + 1])


def snap_to_center() -> None:
    wnd = keymap.getTopLevelWindow()
    wnd_left, wnd_top, wnd_right, wnd_bottom = wnd.getRect()
    to_x = int((wnd_left + wnd_right) / 2)
    to_y = int((wnd_bottom + wnd_top) / 2)
    set_position(to_x, to_y)
