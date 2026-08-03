from . import main, style  # noqa: N999


def configure(keymap) -> None:
    style.setup(keymap)
    main.setup(keymap)
