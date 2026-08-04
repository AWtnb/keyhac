from keyhac import *  # type: ignore

from .tools import cursor_pos as cursor_pos_tool
from .tools.common import is_global_target


def bind(keymap) -> None:
    cursor_pos_tool.setup(keymap)
    km = keymap.defineWindowKeymap(check_func=is_global_target)
    km["O-RCtrl"] = cursor_pos_tool.snap_cursor
    km["O-RShift"] = cursor_pos_tool.snap_to_center
