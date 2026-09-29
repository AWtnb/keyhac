REMAP_SINGLE_KEY = {
    "U1-F": (
        "cfiler.exe",
        "CfilerWindowClass",
        r"${USERPROFILE}\Personal\portable_apps\cfiler\cfiler.exe",
    ),
    "U1-P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
    "U1-K": ("KIRI10.exe", "*"),
    "C-U1-S": ("smoothcsv-app.exe", "*"),
    "LC-U1-M": (
        "Mery.exe",
        "TChildForm",
        r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
    ),
    "LC-U1-N": (
        "notepad.exe",
        "Notepad",
        r"C:\Windows\System32\notepad.exe",
    ),
    "LC-AtMark": (
        "wezterm-gui.exe",
        "org.wezfurlong.wezterm",
        r"${USERPROFILE}\scoop\apps\wezterm\current\wezterm-gui.exe",
    ),
}

REMAP_KEY_SEQUENCE = {
    "Space": (
        SYSTEM_BROWSER.get_exe_name(),
        SYSTEM_BROWSER.get_wnd_class(),
        SYSTEM_BROWSER.get_exe_path(),
    ),
    "C": (
        "chrome.exe",
        "Chrome_WidgetWin_1",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    ),
    "D": (
        "vivaldi.exe",
        "Chrome_WidgetWin_1",
        r"${LOCALAPPDATA}\Vivaldi\Application\vivaldi.exe",
    ),
    "S": (
        "slack.exe",
        "Chrome_WidgetWin_1",
        r"${LOCALAPPDATA}\slack\slack.exe",
    ),
    "F": (
        "firefox.exe",
        "MozillaWindowClass",
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
    ),
    "B": (
        "thunderbird.exe",
        "MozillaWindowClass",
        r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
    ),
    "K": (
        "ksnip.exe",
        "Qt5152QWindowIcon",
        r"${USERPROFILE}\scoop\apps\ksnip\current\ksnip.exe",
    ),
    "O": ("Obsidian.exe", "Chrome_WidgetWin_1"),
    "P": ("SumatraPDF.exe", "SUMATRA_PDF_FRAME"),
    "C-P": ("powerpnt.exe", "PPTFrameClass"),
    "E": ("EXCEL.EXE", "XLMAIN"),
    "W": ("WINWORD.EXE", "OpusApp"),
    "V": ("Code.exe", "Chrome_WidgetWin_1", _open_vscode),
    "C-V": ("vivaldi.exe", "Chrome_WidgetWin_1"),
    "M": (
        "Mery.exe",
        "TChildForm",
        r"${LOCALAPPDATA}\Programs\Mery\Mery.exe",
    ),
    "X": ("explorer.exe", "CabinetWClass", r"C:\Windows\explorer.exe"),
}


def bind(keymap) -> None:
    kt = keymap.define_keytable(focus_path_pattern="*")
