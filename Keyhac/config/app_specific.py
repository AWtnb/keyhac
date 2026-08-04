from .tools import ime as ime_tool
from .tools import sender as sender_tool
from .tools import virtual_finger as vf_tool
from .tools.clipboard import copy_then
from .tools.common import is_browser


def bind_browser(keymap) -> None:
    keymap_browser = keymap.defineWindowKeymap(check_func=is_browser)
    keymap_browser["LC-LS-W"] = "A-Left"
    keymap_browser["LC-F"] = sender_tool.SKKSender(40).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-F"
    )
    keymap_browser["LC-K"] = sender_tool.SKKSender(40).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-K"
    )


def bind_intra(keymap) -> None:
    keymap_intra = keymap.defineWindowKeymap(exe_name="APARClientAWS.exe")
    keymap_intra["O-(235)"] = lambda: None


def bind_slack(keymap) -> None:
    keymap_slack = keymap.defineWindowKeymap(
        exe_name="slack.exe", class_name="Chrome_WidgetWin_1"
    )
    keymap_slack["C-K"] = sender_tool.SKKSender().invoke_emitThen(
        ime_tool.ImeStatus.off, "C-K"
    )
    keymap_slack["F3"] = keymap_slack["C-K"]
    keymap_slack["C-E"] = keymap_slack["C-K"]
    keymap_slack["F1"] = sender_tool.DirectSender().invoke("S-SemiColon", "Colon")


def bind_vscode(keymap) -> None:
    # vscode
    keymap_vscode = keymap.defineWindowKeymap(exe_name="Code.exe")
    keymap_vscode["U0-Slash"] = "C-Slash", "A-S-Down", "C-Slash"

    def remap_vscode(*keys: str) -> None:
        sender = sender_tool.SKKSender()
        for key in keys:
            keymap_vscode[key] = sender.invoke_emitThen(ime_tool.ImeStatus.off, key)

    remap_vscode(
        "C-E",
        "C-F",
        "C-T",
        "C-S-F",
        "C-S-E",
        "C-S-O",
        "C-S-G",
        "RC-RS-X",
        "C-0",
        "C-S-P",
        "C-A-B",
        "C-A-AtMark",
        "C-1",
        "C-2",
        "C-S-Enter",
        "S-Enter",
    )


def bind_mery(keymap) -> None:
    keymap_mery = keymap.defineWindowKeymap(exe_name="Mery.exe")

    def remap_mery(binding: dict) -> None:
        for key, value in binding.items():
            keymap_mery[key] = value

    remap_mery(
        {
            "LA-LC-J": "LA-LC-N",
            "LA-LC-K": "LA-LC-LS-N",
            "LA-U0-J": "A-CloseBracket",
            "LA-U0-K": "A-OpenBracket",
            "LA-LC-U0-J": "A-C-CloseBracket",
            "LA-LC-U0-K": "A-C-OpenBracket",
            "LA-LS-U0-J": "A-S-CloseBracket",
            "LA-LS-U0-K": "A-S-OpenBracket",
        }
    )


def bind_kiri(keymap) -> None:
    keymap_kiri = keymap.defineWindowKeymap(
        exe_name="KIRI10.exe", class_name="TblCommCtrl"
    )
    keymap_kiri["F2"] = "F2", "End"
    keymap_kiri["U0-N"] = keymap_kiri["F2"]

    keymap_kiri_edit = keymap.defineWindowKeymap(
        exe_name="KIRI10.exe", class_name="Edit"
    )
    keymap_kiri_edit["C-Enter"] = "F4", "Down"
    keymap_kiri_edit["LC-U0-Space"] = keymap_kiri_edit["C-Enter"]


def bind_smooth_csv(keymap) -> None:
    keymap_smoothcsv = keymap.defineWindowKeymap(
        exe_name="msedgewebview2.exe",
        class_name="Chrome_WidgetWin_1",
        window_text="tauri.localhost",
    )
    keymap_smoothcsv["C-S-F"] = sender_tool.SKKSender(80).invoke_emitThen(
        ime_tool.ImeStatus.off, "C-S-F", "C-A"
    )
    keymap_smoothcsv["S-Space"] = sender_tool.DirectSender().invoke("S-Space")
    keymap_smoothcsv["S-U0-N"] = lambda: vf_tool.VirtualFinger(20).send("F2", "Home")

    def copy_and_unselect_line() -> None:
        finger = vf_tool.VirtualFinger()
        taps = vf_tool.VirtualFinger().compile("Up", "Down")

        def _unselect(_) -> None:
            finger.send_compiled(*taps)

        copy_then(_unselect)

    keymap_smoothcsv["U1-C"] = copy_and_unselect_line


def bind_sumatra_pdf(keymap) -> None:
    keymap_sumatra_view = keymap.defineWindowKeymap(
        check_func=(
            lambda wnd: (
                wnd.getProcessName() == "SumatraPDF.exe"
                and wnd.getClassName() != "Edit"
            )
        )
    )

    sender = sender_tool.DirectSender()
    for key in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        keymap_sumatra_view[key] = sender.invoke(key)


def bind_office_powerpoint(keymap) -> None:
    keymap_ppt = keymap.defineWindowKeymap(exe_name="powerpnt.exe")
    keymap_ppt["O-(236)"] = ime_tool.Handler(40).to_skk_abbrev


def bind_office_excel(keymap) -> None:
    keymap_excel = keymap.defineWindowKeymap(exe_name="excel.exe")

    def select_all() -> None:
        finger = vf_tool.VirtualFinger()
        if keymap.getWindow().getClassName() == "EXCEL6":
            finger.send("C-End", "C-S-Home")
        else:
            finger.send("C-A")

    keymap_excel["C-A"] = select_all


def setup(keymap) -> None:
    sender_tool.setup(keymap)
    ime_tool.setup(keymap)
    vf_tool.setup(keymap)

    bind_browser(keymap)
    bind_intra(keymap)
    bind_slack(keymap)
    bind_vscode(keymap)
    bind_mery(keymap)
    bind_kiri(keymap)
    bind_smooth_csv(keymap)
    bind_sumatra_pdf(keymap)
    bind_office_powerpoint(keymap)
    bind_office_excel(keymap)
