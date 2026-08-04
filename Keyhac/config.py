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
    config = setup_config("config_listwindow")
    config.configure(window)
