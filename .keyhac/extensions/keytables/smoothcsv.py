from keyhac import ThreadedAction  # ty: ignore[unresolved-import]
from libs import clipboard, sender
from libs._common import delay


def bind(keymap) -> None:
    sender.setup(keymap)
    clipboard.setup(keymap)

    kt = keymap.define_keytable(app="smoothcsv-app")
    kt["S-Space"] = sender.DirectSender().invoke("S-Space")

    def _unselect(_) -> None:
        sender.send_keys("Up", "Down")

    kt["U1-C"] = clipboard.CopyThen(_unselect)

    class LazyFilter(ThreadedAction):
        def starting(self):
            sender.send_keys("C-S-F")

        def run(self):
            delay(100)

        def finished(self, _):
            sender.send_keys("C-A")

    kt["LC-LS-F"] = LazyFilter()
