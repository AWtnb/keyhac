import re
import urllib.parse
from collections.abc import Callable

from libs import clipboard
from libs.web_search import cleanup_web_search_query

REG_HIRAGANA = re.compile(r"[\u3041-\u3093]")
REG_SPACES = re.compile(r"[ 　]+")


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


def bind(keymap) -> None:
    clipboard.setup(keymap)

    def build_web_searcher(
        uri: str, strict: bool = False, omit_hiragana: bool = False
    ) -> Callable[[str], None]:

        def _searcher(s: str) -> None:
            query = cleanup_web_search_query(s)
            if omit_hiragana:
                query = REG_HIRAGANA.sub(" ", query)

            words = []
            for word in REG_SPACES.split(query):
                if strict:
                    words.append(f'"{word}"')
                else:
                    words.append(word)
            url = uri.format(urllib.parse.quote(" ".join(words)))
            keymap.app_control.open_url(url)

        return _searcher

    kt = keymap.define_keytable(focus_path_pattern="*")

    for shift_key in ("", "S-"):
        for ctrl_key in ("", "C-"):
            strict_frag = shift_key != ""
            hiragana_omit_flag = ctrl_key != ""
            kt_search = keymap.define_keytable(
                name=f"search (strict={strict_frag!r}, omit_hiragana={hiragana_omit_flag!r})"
            )

            for key, uri in MAPPING.items():
                searcher = build_web_searcher(
                    uri=uri,
                    strict=strict_frag,
                    omit_hiragana=hiragana_omit_flag,
                )
                kt_search[key] = clipboard.CopyThen(searcher)

            trigger_key = shift_key + ctrl_key + "U0-S"
            kt[trigger_key] = kt_search
