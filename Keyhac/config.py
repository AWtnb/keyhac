import importlib
import sys
from pathlib import Path

from keyhac_listwindow import ListWindow  # type: ignore

from keyhac import *  # type: ignore

CONFIG_MODULE_NAME = "config"


def configure(keymap) -> None:

    config_dir = str(Path(keymap.config_filename).parent)
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    for name in list(sys.modules):
        if name == CONFIG_MODULE_NAME or name.startswith(CONFIG_MODULE_NAME + "."):
            del sys.modules[name]
    config = importlib.import_module(CONFIG_MODULE_NAME)
    config.configure(keymap)


def configure_ListWindow(window: ListWindow) -> None:
    window.keymap["J"] = window.command_CursorDown
    window.keymap["K"] = window.command_CursorUp
    window.keymap["C-J"] = window.command_CursorPageDown
    window.keymap["C-K"] = window.command_CursorPageUp
    window.keymap["L"] = window.command_Enter
    for mod in ["", "S-", "C-", "C-S-"]:
        for key in ["L", "Space"]:
            window.keymap[mod + key] = window.command_Enter

    def to_top_of_list() -> None:
        if window.isearch:
            return
        window.select = 0
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["A"] = to_top_of_list

    def to_end_of_list() -> None:
        if window.isearch:
            return
        window.select = len(window.items) - 1
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["E"] = to_end_of_list
