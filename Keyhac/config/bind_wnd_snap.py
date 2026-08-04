from keyhac import *  # type: ignore

from .tools import window_snap as wnd_snap_tool
from .tools.common import is_global_target
from .tools.window_rect import RectEdge


def bind(keymap) -> None:

    km = keymap.defineWindowKeymap(check_func=is_global_target)
    wnd_snap_tool.setup(keymap)

    ################################
    # set window position
    ################################

    km["U1-L"] = "LWin-Right"
    km["U1-H"] = "LWin-Left"

    km["U1-M"] = keymap.defineMultiStrokeKeymap()
    km["U1-M"]["X"] = lambda: keymap.getTopLevelWindow().maximize()
    km["U1-M"]["N"] = lambda: keymap.getTopLevelWindow().minimize()

    altkey_stat = {0: "", 1: "LA-", 2: "RA-"}
    scale_mapping = {
        "": 1 / 2,
        "S-": 2 / 3,
        "C-": 1 / 3,
    }
    edge_mapping = {
        "H": RectEdge.left,
        "L": RectEdge.right,
        "J": RectEdge.bottom,
        "K": RectEdge.top,
    }
    for idx, alt in altkey_stat.items():
        for area_mod, scale in scale_mapping.items():
            for key, edge in edge_mapping.items():
                km["U1-M"][alt + area_mod + key] = wnd_snap_tool.invoke_snapper(
                    idx, scale, edge
                )

    for key in ["0", "1", "2"]:
        monitor_idx = int(key)
        km["U1-M"][key] = wnd_snap_tool.invoke_maximized_snapper(monitor_idx)

    for key, toward in {
        "H": RectEdge.left,
        "L": RectEdge.right,
        "K": RectEdge.top,
        "J": RectEdge.bottom,
    }.items():
        km["U1-M"]["U1-" + key] = wnd_snap_tool.invoke_shrinker(toward)
