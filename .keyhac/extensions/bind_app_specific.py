from keyhac import ThreadedAction  # ty: ignore[unresolved-import]
from libs import clipboard, sender
from libs._common import delay


def _bind_browser(keymap) -> None:
    kt = keymap.define_keytable(app="chrome|Google Chrome|firefox|Safari")
    kt["LC-LS-W"] = "A-Left"

    pause = 40

    skk_sender = sender.SKKSender(pause)
    for k in ["F", "K"]:
        key = f"LC-{k}"
        kt[key] = skk_sender.invoke_emitThen(False, key)

    def slow_bookmark() -> None:
        with keymap.get_input_context() as ctx:
            delay(pause)
            ctx.send_key("C-D")

    kt["LC-D"] = slow_bookmark


def _bind_slack(keymap) -> None:
    kt = keymap.define_keytable(app="slack.exe", class_name="Chrome_WidgetWin_1")
    kt["C-K"] = sender.SKKSender().invoke_emitThen(False, "C-K")
    kt["F3"] = kt["C-K"]
    kt["C-E"] = kt["C-K"]
    kt["F1"] = sender.DirectSender().invoke("S-SemiColon", "Colon")


def _bind_vscode(keymap) -> None:
    kt = keymap.define_keytable(app="Code.exe")
    kt["U0-Slash"] = "C-Slash", "A-S-Down", "C-Slash"

    for key in [
        "C-E",
        "C-F",
        "C-T",
        "C-S-F",
        "C-S-E",
        "C-S-O",
        "C-S-G",
        "RC-RS-X",
        "C-0",
        "C-S-P",
        "C-A-B",
        "C-A-AtMark",
        "C-1",
        "C-2",
        "C-S-Enter",
        "S-Enter",
    ]:
        kt[key] = sender.SKKSender().invoke_emitThen(False, key)


def _bind_mery(keymap) -> None:
    kt = keymap.define_keytable(app="Mery.exe")

    for key, value in {
        "LA-LC-J": "LA-LC-N",
        "LA-LC-K": "LA-LC-LS-N",
        "LA-U0-J": "A-CloseBracket",
        "LA-U0-K": "A-OpenBracket",
        "LA-LC-U0-J": "A-C-CloseBracket",
        "LA-LC-U0-K": "A-C-OpenBracket",
        "LA-LS-U0-J": "A-S-CloseBracket",
        "LA-LS-U0-K": "A-S-OpenBracket",
    }.items():
        kt[key] = value


def _bind_smooth_csv(keymap) -> None:
    kt = keymap.define_keytable(
        focus_path_pattern="/Application(smoothcsv-app)/Window(SmoothCSV)/*"
    )
    kt["S-Space"] = sender.DirectSender().invoke("S-Space")

    def _unselect(_) -> None:
        with keymap.get_input_context() as ctx:
            for key in ("Up", "Down"):
                ctx.send_key(key)

    kt["U1-C"] = clipboard.CopyThen(_unselect)

    class LazyFilter(ThreadedAction):
        def starting(self):
            with keymap.get_input_context() as ctx:
                ctx.send_key("C-S-F")

        def run(self):
            delay(100)

        def finished(self, _):
            with keymap.get_input_context() as ctx:
                ctx.send_key("C-A")

    kt["LC-LS-F"] = LazyFilter()


def _bind_sumatra_pdf(keymap) -> None:
    kt = keymap.define_keytable(
        custom_condition_func=(
            lambda focus: focus.app_name == "SumatraPDF" and focus.class_name != "Edit"
        )
    )

    for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        kt[key] = sender.DirectSender().invoke(key)


def _bind_office_excel(keymap) -> None:
    kt = keymap.define_keytable(app="excel.exe")

    def select_all() -> None:
        focus = keymap.focus
        if focus is None:
            return

        sequence = ["C-End", "C-S-Home"] if focus.class_name == "EXCEL6" else ["C-A"]
        with keymap.get_input_context() as ctx:
            for s in sequence:
                ctx.send_key(s)

    kt["C-A"] = select_all


def bind(keymap):
    sender.setup(keymap)
    clipboard.setup(keymap)

    _bind_browser(keymap)
    _bind_slack(keymap)
    _bind_vscode(keymap)
    _bind_mery(keymap)
    _bind_smooth_csv(keymap)
    _bind_sumatra_pdf(keymap)
    _bind_office_excel(keymap)
