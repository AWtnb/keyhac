from libs import sender
from libs._common import delay


def bind_browser(keymap) -> None:
    kt = keymap.define_keytable(app="chrome|Google Chrome|firefox|Safari")
    kt["LC-LS-W"] = "A-Left"

    pause = 40

    skk_sender = sender.SKKSender(pause)
    for k in ["F", "K"]:
        key = f"LC-{k}"
        kt[key] = skk_sender.invoke_emitThen(False, key)

    def slow_bookmark() -> None:
        delay()
        sender.send_sequence(pause, "C-D")

    kt["LC-D"] = slow_bookmark


def bind(keymap):
    sender.setup(keymap)
    bind_browser(keymap)
