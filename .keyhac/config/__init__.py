from . import bind_core  # noqa: N999


def configure(keymap) -> None:
    bind_core.bind(keymap)
