import os
from winreg import HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, OpenKey, QueryValueEx


class SystemBrowser:
    prog_id: str = ""
    commandline: str = ""

    def __init__(self) -> None:
        self.set_prog_id()
        self.set_commandline()

    def set_prog_id(self) -> None:
        registry_paths = [
            r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoiceLatest\ProgId",
            r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice",
        ]

        for path in registry_paths:
            try:
                with OpenKey(HKEY_CURRENT_USER, path) as key:
                    self.prog_id = str(QueryValueEx(key, "ProgId")[0])
                    return
            except Exception as e:  # noqa: BLE001
                print(e)
                print(f"Failed to get ProgId by registry `{path}`")
                if path != registry_paths[-1]:
                    print("Try next path...")

    def set_commandline(self) -> None:
        if self.prog_id == "":
            return
        registry_path = os.path.join(self.prog_id, "shell", "open", "command")
        try:
            with OpenKey(HKEY_CLASSES_ROOT, registry_path) as key:
                self.commandline = str(QueryValueEx(key, "")[0])
        except Exception as e:  # noqa: BLE001
            print(e)
            print(f"Failed to get ProgId by registry `{registry_path}`")

    def get_exe_path(self) -> str:
        if self.commandline == "" or self.prog_id == "":
            return ""
        c = self.commandline
        e = ".exe"
        return c[: c.find(e) + len(e)].strip('"')

    def get_exe_name(self) -> str:
        if self.commandline == "" or self.prog_id == "":
            return ""
        _, name = os.path.split(self.get_exe_path())
        return name

    def get_wnd_class(self) -> str:
        return {
            "chrome.exe": "Chrome_WidgetWin_1",
            "vivaldi.exe": "Chrome_WidgetWin_1",
            "firefox.exe": "MozillaWindowClass",
        }.get(self.get_exe_name(), "Chrome_WidgetWin_1")
