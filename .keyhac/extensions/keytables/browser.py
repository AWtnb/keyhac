from keyhac import ThreadedAction  # ty: ignore[unresolved-import]
from libs import ime, sender
from libs._common import delay


def bind(keymap) -> None:
    sender.setup(keymap)
    ime.setup(keymap)

    kt = keymap.define_keytable(app="chrome|Google Chrome|firefox|Safari")

    kt["LC-LS-W"] = "A-Left"

    for k in ["F", "K", "L"]:
        key = f"LC-{k}"
        kt[key] = sender.SendThen(ime.turnoff_skk, [key], 40)

    class LazyBookmark(ThreadedAction):
        def run(self):
            delay(200)

        def finished(self, _):
            sender.send_keys("C-D")

    kt["LC-D"] = LazyBookmark()
