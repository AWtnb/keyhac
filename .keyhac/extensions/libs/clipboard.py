def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


def get_string() -> str:
    return keymap.clipboard.get_text() or ""


def set_string(s: str) -> None:
    keymap.clipboard.set_text(s)


def _send_key(key: str) -> None:
    with keymap.get_input_context() as ctx:
        ctx.send_key(key)


def send_copy_key() -> None:
    _send_key("C-C")


def send_paste_key() -> None:
    _send_key("C-V")
