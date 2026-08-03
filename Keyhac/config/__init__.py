from . import app_based, main, style  # noqa: N999


def configure(keymap) -> None:
    style.setup(keymap)
    main.setup(keymap)
    app_based.setup(keymap)
