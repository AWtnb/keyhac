import os
import shutil
import subprocess
import time

from keyhac import (  # ty: ignore[unresolved-import]
    ChooserPage,
    ClipboardHistorySource,
    MouseHorizontalWheel,
    MouseWheel,
    MoveWindow,
    PlaybackRecordedKeys,
    ShowCandidates,
    ThreadedAction,
    ToggleRecordingKeys,
)


def bind(keymap) -> None:
    kt = keymap.define_keytable(focus_path_pattern="*")

    def clipboard_history():
        chooser_action = ShowCandidates(
            [ChooserPage("Clipboard History", ClipboardHistorySource())]
        )
        chooser_action.activates = True
        chooser_action()

    kt["LC-LS-X"] = clipboard_history

    # keyboard macro
    kt["U0-0"] = ToggleRecordingKeys()
    kt["U1-0"] = PlaybackRecordedKeys()
    kt["U1-F4"] = PlaybackRecordedKeys()

    # mouse scroll
    kt["U1-Up"] = MouseWheel(1.0)
    kt["U1-Down"] = MouseWheel(-1.0)
    kt["U1-Left"] = MouseHorizontalWheel(-1.0)
    kt["U1-Right"] = MouseHorizontalWheel(1.0)

    # window mover
    window_move_unit = 10  # px
    for key in ["Left", "Right", "Up", "Down"]:
        direction = key.lower()
        for mod, scale in {"": 15, "S-": 5, "C-": 5, "S-C-": 1}.items():
            kt[f"{mod}U0-{key}"] = MoveWindow(
                direction=direction, distance=window_move_unit * scale
            )

    mod_keys = ("", "S-", "C-", "A-", "C-S-", "C-A-", "S-A-", "C-A-S-")
    for mod_key in mod_keys:
        for key, value in {
            # move cursor
            "H": "Left",
            "J": "Down",
            "K": "Up",
            "L": "Right",
            # Back / Delete
            "B": "Back",
            "D": "Delete",
            # Home / End
            "A": "Home",
            "E": "End",
            # Enter
            "Space": "Enter",
        }.items():
            kt[f"{mod_key}U0-{key}"] = mod_key + value

    for key, value in {
        # delete to bol / eol
        "S-U0-B": ("S-Home", "Delete"),
        "S-U0-D": ("S-End", "Delete"),
        # escape
        "O-(235)": ("Esc"),
        "U0-X": ("Esc"),
        # line selection
        "U1-A": ("End", "S-Home"),
        # punctuation
        "U0-U": ("S-BackSlash"),
        "U1-S": ("Slash"),
        "U0-4": ("S-4", "S-BackSlash"),
        "U0-Enter": ("Period"),
        "U0-Z": ("Minus"),
        "U1-X": ("S-1"),
        # Insert line
        "U0-I": ("End", "Enter"),
        "S-U0-I": ("Home", "Enter", "Up"),
        # Context menu
        "U0-C": ("Apps"),
        "S-U0-C": ("S-Apps"),
        # rename
        "U0-N": ("F2"),
    }.items():
        kt[key] = value

    for key, value in {
        "U0-2": "LS-2",
        "U0-7": "LS-7",
    }.items():
        kt[key] = value, value, "Left"

    class LazyReload(ThreadedAction):
        def run(self) -> None:
            keymap.configure()

    kt["U1-F12"] = LazyReload()

    class CarefulQuit(ThreadedAction):
        def run(self) -> None:
            time.sleep(0.25)

        def finished(self, _) -> None:
            with keymap.get_input_context() as ctx:
                ctx.send_key("A-F4")

    kt["LC-Q"] = CarefulQuit()

    class ConfigRepoLauncher(ThreadedAction):
        def run(self) -> None:
            config_path = os.path.expandvars(r"${USERPROFILE}\.keyhac")
            if not os.path.exists(config_path):
                print(f"config not found: {config_path}")
                return

            dir_path = config_path
            if (real_path := os.path.realpath(config_path)) != dir_path:
                dir_path = os.path.dirname(real_path)

            vscode_path = shutil.which("code")

            cmd = ["explorer.exe", dir_path]
            if vscode_path is not None:
                cmd[0] = vscode_path

            try:
                subprocess.run(
                    cmd, creationflags=subprocess.CREATE_NO_WINDOW, check=False
                )
            except Exception as e:  # noqa: BLE001
                print(e)

    config_launcher = ConfigRepoLauncher()
    keymap.editor = lambda _: config_launcher()

    kt["U0-F12"] = config_launcher
