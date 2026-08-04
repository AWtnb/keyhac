import importlib
import sys
from types import ModuleType

import ckit  # type: ignore
from keyhac_listwindow import ListWindow  # type: ignore

from keyhac import *  # type: ignore


def setup_config(config_module_name: str) -> ModuleType:
    config_dir = ckit.dataPath()
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    for name in list(sys.modules):
        if name == config_module_name or name.startswith(config_module_name + "."):
            del sys.modules[name]
    return importlib.import_module(config_module_name)


def configure(keymap) -> None:
    config = setup_config("config")
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
