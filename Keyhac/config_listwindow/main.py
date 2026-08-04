def setup(window) -> None:
    window.keymap["J"] = window.command_CursorDown
    window.keymap["K"] = window.command_CursorUp
    window.keymap["C-J"] = window.command_CursorPageDown
    window.keymap["C-K"] = window.command_CursorPageUp
    window.keymap["L"] = window.command_Enter
    for mod in ["", "S-", "C-", "C-S-"]:
        for key in ["L", "Space"]:
            window.keymap[mod + key] = window.command_Enter

    def to_top_of_list() -> None:
        if window.isearch:
            return
        window.select = 0
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["A"] = to_top_of_list

    def to_end_of_list() -> None:
        if window.isearch:
            return
        window.select = len(window.items) - 1
        window.scroll_info.makeVisible(window.select, window.itemsHeight())
        window.paint()

    window.keymap["E"] = to_end_of_list
