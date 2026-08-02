import ckit  # type: ignore
import keyhac_ini  # type: ignore


def set_custom_theme(keymap) -> None:
    name = "black"

    custom_theme = {
        "bg": "#3f3b39",
        "fg": "#a0b4a7",
        "cursor0": "#ffffff",
        "cursor1": "#ff4040",
        "bar_fg": "#000000",
        "bar_error_fg": "#ff4040",
        "select_bg": "#dff477",
        "select_fg": "#3f3b39",
        "caret0": "#ffffff",
        "caret1": "#ff0000",
    }
    ckit.ckit_theme.theme_name = name

    for k, v in custom_theme.items():
        rgb = tuple(int(v[i : i + 2], 16) for i in (1, 3, 5))
        ckit.ckit_theme.ini.set("COLOR", k, str(rgb))
    keymap.console_window.reloadTheme()


def setup(keymap) -> None:
    keymap.setFont("HackGen", 16)

    set_custom_theme(keymap)

    # set console appearance
    keyhac_ini.setint("CONSOLE", "visible", 0)
    keyhac_ini.setint("CONSOLE", "x", 0)
    keyhac_ini.write()
