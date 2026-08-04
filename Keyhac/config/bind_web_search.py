import ckit  # type: ignore

from keyhac import *  # type: ignore

from .tools.clipboard import copy_then
from .tools.common import CallbackFunc, is_global_target
from .tools.web_search import invoke_web_seacher


def invoke_copied_str_searcher(
    uri: str, strict: bool, strip_hiragana: bool
) -> CallbackFunc:
    func = invoke_web_seacher(uri, strict, strip_hiragana)

    def _search(job_item: ckit.JobItem) -> None:
        s = job_item.copied
        if len(s) < 1:
            s = job_item.origin
        func(s)

    def _searcher() -> None:
        copy_then(_search)

    return _searcher


def bind(keymap) -> None:
    km = keymap.defineWindowKeymap(check_func=is_global_target)

    MAPPING = {
        "A": "https://www.amazon.co.jp/s?i=stripbooks&k={}",
        "B": "https://www.google.com/search?nfpr=1&q=site%3Abooks.or.jp%20{}",
        "C": "https://ci.nii.ac.jp/books/search?q={}",
        "D": "https://duckduckgo.com/?q={}",
        "G": "http://www.google.com/search?nfpr=1&q={}",
        "H": "https://www.hanmoto.com/bd/search/order/desc/title/{}",
        "I": "https://www.google.com/search?udm=2&nfpr=1&q={}",
        "J": "https://eow.alc.co.jp/search?q={}",
        "M": "https://www.merriam-webster.com/dictionary/{}",
        "N": "https://ndlsearch.ndl.go.jp/search?cs=bib&f-ht=ndl&keyword={}",
        "P": "https://wordpress.org/openverse/search/?q={}",
        "R": "https://researchmap.jp/researchers?q={}",
        "S": "https://scholar.google.com/scholar?nfpr=1&as_vis=1&q={}",
        "T": "https://twitter.com/search?q={}",
        "Y": "https://duckduckgo.com/?q=site%3Ayuhikaku.co.jp%20{}",
        "W": "https://www.worldcat.org/search?q={}",
    }

    for shift_key in ("", "S-"):
        for ctrl_key in ("", "C-"):
            is_strict = shift_key != ""
            strip_hiragana = ctrl_key != ""
            trigger_key = shift_key + ctrl_key + "U0-S"
            km[trigger_key] = keymap.defineMultiStrokeKeymap()

            for key, uri in MAPPING.items():
                km[trigger_key][key] = invoke_copied_str_searcher(
                    uri, is_strict, strip_hiragana
                )
