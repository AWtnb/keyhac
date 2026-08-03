from . import app_specific, main, style  # noqa: N999


def configure(keymap) -> None:
    style.setup(keymap)
    main.setup(keymap)
    app_specific.setup(keymap)
