import os

import ckit  # type: ignore

from keyhac import *  # type: ignore

from .tools import subthread as subthread_tool
from .tools import virtual_finger as vf_tool
from .tools.common import (
    balloon,
    delay,
    get_now,
    is_global_target,
    open_vscode,
    shell_exec,
    smart_check_path,
)
from .tools.virtual_finger import Tap


def bind(keymap) -> None:

    vf_tool.setup(keymap)
    subthread_tool.setup(keymap)

    km = keymap.defineWindowKeymap(check_func=is_global_target)

    VF = vf_tool.VirtualFinger(0)

    def safe_close() -> None:
        close_tap = Tap("A-F4")

        def _wait(_) -> None:
            delay(200)

        def _close(_) -> None:
            VF.send_compiled(close_tap)

        subthread_tool.run(_wait, _close)

    km["C-Q"] = safe_close

    def reload_config() -> None:
        def _wait(_) -> None:
            delay(60)

        def _reload(_) -> None:
            ckit.JobQueue.cancelAll()
            keymap.configure()
            keymap.updateKeymap()
            ts = get_now().strftime("%Y-%m-%d %H:%M:%S")
            balloon(keymap, f"{ts} reloaded config.py")

        subthread_tool.run(_wait, _reload)

    km["U1-F12"] = reload_config

    def open_keyhac_repo() -> None:
        config_path = os.path.expandvars(r"${APPDATA}\Keyhac")
        if not smart_check_path(config_path):
            balloon(keymap, f"config not found: {config_path}")
            return

        dir_path = config_path
        if (real_path := os.path.realpath(config_path)) != dir_path:
            dir_path = os.path.dirname(real_path)

        def _open(_) -> None:
            result = open_vscode(dir_path)
            if not result:
                shell_exec(dir_path)

        subthread_tool.run(_open)

    keymap.editor = lambda _: open_keyhac_repo()

    km["U0-F12"] = open_keyhac_repo
