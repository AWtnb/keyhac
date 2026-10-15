from libs import ime, sender


def bind(keymap) -> None:
    sender.setup(keymap)
    ime.setup(keymap)

    kt = keymap.define_keytable(app="slack.exe", class_name="Chrome_WidgetWin_1")

    kt["C-K"] = sender.SendThen(ime.turnoff_skk, ["C-K"], 60)
    kt["F3"] = kt["C-K"]
    kt["C-E"] = kt["C-K"]
    kt["F1"] = sender.DirectSender().invoke("S-SemiColon", "Colon")
    kt["C-O"] = sender.DirectSender().invoke("C-O")
