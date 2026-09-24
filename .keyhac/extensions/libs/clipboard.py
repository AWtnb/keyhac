def setup(_keymap) -> None:
    global keymap  # ty: ignore[unresolved-global]
    keymap = _keymap


def get_string() -> str:
    return keymap.clipboard.get_text() or ""


def get_latest_clipboard_history() -> str:
    return keymap.clipboard_history.get_current() or ""


def set_string(s: str) -> None:
    keymap.clipboard.set_text(s)


def _send_key(key: str) -> None:
    with keymap.get_input_context() as ctx:
        ctx.send_key(key)


def send_copy_key() -> None:
    _send_key("C-C")


def send_paste_key() -> None:
    _send_key("C-V")

    # # ---- Background work: ThreadedAction -----------------------------

    # # Anything slow (network, subprocess, sleeping) must not run inline - it
    # # would block the keyboard hook.  run() is on a worker thread; starting()
    # # and finished() stay on the main thread, so UI and window access is fine
    # # in those two and not in run().  Esc stops a running action.
    # class TypeSlowly(ThreadedAction):
    #     def __init__(self, text):
    #         self.text = text

    #     def starting(self):
    #         logger.info(f"Typing {self.text!r}...")

    #     def run(self):
    #         import time
    #         for char in self.text:
    #             time.sleep(0.05)
    #             with keymap.get_input_context() as ctx:
    #                 ctx.send_key(f"Shift-{char}" if char.isupper() else char)
    #         return len(self.text)

    #     def finished(self, result):
    #         logger.info(f"Typed {result} characters.")
