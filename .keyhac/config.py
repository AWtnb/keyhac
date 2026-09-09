import importlib
import os
import sys
from types import ModuleType


def import_config(config_module_name: str) -> ModuleType:
    user_profile = os.environ.get("USERPROFILE")
    assert user_profile is not None
    config_dir = os.path.join(user_profile, ".keyhac")
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    for name in list(sys.modules):
        if name == config_module_name or name.startswith(config_module_name + "."):
            del sys.modules[name]

    return importlib.import_module(config_module_name)


def configure(keymap) -> None:
    config = import_config("config")
    config.configure(keymap)
