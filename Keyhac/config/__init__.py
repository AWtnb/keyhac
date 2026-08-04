from . import app_specific, main, style  # noqa: N999


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
    style.setup(keymap)
    main.setup(keymap)
    app_specific.setup(keymap)
