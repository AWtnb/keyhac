import importlib
import os
import sys
from types import ModuleType


def register_config_path(dir_name: str) -> None:
    user_profile = os.environ.get("USERPROFILE")
    assert user_profile is not None
    config_dir = os.path.join(user_profile, dir_name)
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)


register_config_path(".keyhac")


def import_config(config_module_name: str) -> ModuleType:
    for name in list(sys.modules):
        if name == config_module_name or name.startswith(config_module_name + "."):
            del sys.modules[name]

    return importlib.import_module(config_module_name)


def configure(keymap) -> None:
    config = import_config("config")
    config.configure(keymap)
