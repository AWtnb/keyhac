from collections.abc import Callable

import ckit  # type: ignore

from keyhac import *  # type: ignore

from .tools.clipboard import copy_then
from .tools.common import is_global_target
from .tools.web_search import invoke_web_seacher


def as_job(searcher: Callable[[str], None]) -> Callable[[ckit.JobItem], None]:

    def _search_job(job_item: ckit.JobItem) -> None:
        s = job_item.copied
        if len(s) < 1:
            s = job_item.origin
        searcher(s)

    return _search_job


def bind_search_key(
    keymap, key: str, search_job: Callable[[ckit.JobItem], None]
) -> None:

    def _searcher() -> None:
        copy_then(search_job)

    keymap[key] = _searcher


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
            trigger_key = shift_key + ctrl_key + "U0-S"
            km[trigger_key] = keymap.defineMultiStrokeKeymap()

            for key, uri in MAPPING.items():
                searcher = invoke_web_seacher(
                    uri=uri,
                    strict=shift_key != "",
                    strip_hiragana=ctrl_key != "",
                )
                bind_search_key(km[trigger_key], key, as_job(searcher))
