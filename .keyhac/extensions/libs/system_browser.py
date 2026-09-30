import os
from winreg import HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, OpenKey, QueryValueEx

REGISTRY_PATHS = [
    r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoiceLatest\ProgId",
    r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice",
]


def _get_prog_id() -> str | None:
    for path in REGISTRY_PATHS:
        try:
            with OpenKey(HKEY_CURRENT_USER, path) as key:
                return str(QueryValueEx(key, "ProgId")[0])
        except Exception:  # noqa: BLE001, S110
            pass
    return None


def _get_commandline() -> str | None:
    prog_id = _get_prog_id()
    if not prog_id:
        return None
    registry_path = os.path.join(prog_id, "shell", "open", "command")
    try:
        with OpenKey(HKEY_CLASSES_ROOT, registry_path) as key:
            return str(QueryValueEx(key, "")[0])
    except Exception:  # noqa: BLE001
        return None


def get_browser_path() -> str | None:
    command_line = _get_commandline()
    if command_line is None:
        return None
    ext = ".exe"
    return command_line[: command_line.find(ext) + len(ext)].strip('"')
