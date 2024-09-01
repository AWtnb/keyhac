import sys
sys.dont_write_bytecode = True

import importlib
from setting import setup


def configure(keymap) -> None:
    importlib.reload(setup)
    setup.configure(keymap)
