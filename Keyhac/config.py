import importlib
import sys
from pathlib import Path

from keyhac import *  # type: ignore  # noqa: F403

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
