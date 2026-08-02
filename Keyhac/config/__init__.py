from . import main, style


def configure(keymap) -> None:
    style.setup(keymap)
    main.setup(keymap)
