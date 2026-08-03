import pyauto  # type: ignore
from keyhac_keymap import WindowKeymap  # type: ignore

from keyhac import *  # type: ignore

from . import subthread
from .common import CallbackFunc, delay, get_monitor_infos, is_keyhac_console
from .window_rect import RectEdge, as_rect


def setup(_keymap: WindowKeymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap
    subthread.setup(keymap)


def snap_window(dest: tuple[int, int, int, int]) -> pyauto.Window | None:
    wnd = keymap.getTopLevelWindow()
    if not wnd or is_keyhac_console(wnd):
        return None
    rect = list(dest)
    if wnd.getRect() == rect:
        return None

    def _job_snap(_) -> None:
        if wnd.isMaximized():
            wnd.restore()
            delay()
        wnd.setRect(rect)

    def _job_finished(_) -> None:
        if wnd.getRect() != rect:
            wnd.setRect(rect)

    subthread.run(_job_snap, _job_finished)
    return wnd


def invoke_snapper(monitor_index: int, scale: float, edge: RectEdge) -> CallbackFunc:
    def _snap() -> None:
        infos = get_monitor_infos()
        target = infos[monitor_index]
        monitor_work_rect = as_rect(*target[1])
        dest = monitor_work_rect.resize(scale, edge)
        _ = snap_window(dest)

    return _snap


def invoke_maximized_snapper(monitor_index: int) -> CallbackFunc:
    def _snap() -> None:
        infos = get_monitor_infos()
        try:
            target = infos[monitor_index][1]
        except IndexError:
            return
        wnd = snap_window(target)
        if wnd is not None:
            wnd.maximize()

    return _snap


def invoke_shrinker(toward: RectEdge):
    def _shrink() -> None:
        wnd = keymap.getTopLevelWindow()
        rect = wnd.getRect()
        resized = as_rect(*rect).resize(0.5, toward)
        snap_window(resized)

    return _shrink
