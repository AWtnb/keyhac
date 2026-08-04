from . import (  # noqa: N999
    app_specific,
    bind_clipboard,
    bind_core,
    bind_cursor_snap,
    bind_ime,
    bind_input,
    bind_web_search,
    bind_wnd_activate,
    bind_wnd_snap,
    main,
    style,
)


def configure(keymap) -> None:

    # user modifier
    keymap.replaceKey("(29)", 235)  # "muhenkan" => 235
    keymap.replaceKey("(28)", 236)  # "henkan" => 236
    keymap.defineModifier(235, "User0")  # "muhenkan" => "U0"
    keymap.defineModifier(236, "User1")  # "henkan" => "U1"

    # enable clipbard history
    keymap.clipboard_history.enableHook(True)

    # history max size
    keymap.clipboard_history.maxnum = 200
    keymap.clipboard_history.quota = 10 * 1024 * 1024

    # quote mark when paste with Ctrl.
    keymap.quote_mark = "> "

    # load settings
    bind_core.bind(keymap)
    bind_clipboard.bind(keymap)
    bind_ime.bind(keymap)
    bind_input.bind(keymap)
    bind_wnd_snap.bind(keymap)
    bind_wnd_activate.bind(keymap)
    bind_cursor_snap.bind(keymap)
    bind_web_search.bind(keymap)
    style.setup(keymap)
    main.setup(keymap)
    app_specific.setup(keymap)
