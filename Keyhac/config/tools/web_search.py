import re
import urllib.parse
from collections.abc import Callable

from .common import shell_exec
from .punctuation import KANGXI_RADICAL_MAPPING, RADICAL_MAPPING, SEACH_NOISE_MAPPING


def join_lines(lines: list[str]) -> str:
    def _format(line: str) -> str:
        if line.endswith("-"):
            return line.rstrip("-")
        if len(line.strip()):
            if line[-1].encode("utf-8").isalnum():
                return line + " "
            return line.rstrip()
        return ""

    return "".join([_format(l) for l in lines])


TRANSLATE_TABLE = str.maketrans(
    RADICAL_MAPPING | KANGXI_RADICAL_MAPPING | SEACH_NOISE_MAPPING
)


def cleanup_web_search_query(s: str) -> str:
    lines = (
        s.strip()
        .replace("\u200b", "")
        .replace("\u3000", " ")
        .replace("\t", " ")
        .splitlines()
    )
    query = join_lines(lines).translate(TRANSLATE_TABLE)

    for honor in ["先生", "様"]:
        query = query.replace(honor, " ")

    for honor in [
        "監修",
        "共著",
        "共編著",
        "編著",
        "共編",
        "分担執筆",
        "et al.",
    ]:
        query = query.replace(honor, " ")

    return query


REG_HIRAGANA = re.compile(r"[\u3041-\u3093]")
REG_SPACES = re.compile(r"[ 　]+")


def invoke_web_seacher(
    uri: str, strict: bool = False, strip_hiragana: bool = False
) -> Callable[[str], None]:

    def _searcher(s: str) -> None:
        query = cleanup_web_search_query(s)
        if strip_hiragana:
            query = REG_HIRAGANA.sub(" ", query)

        words = []
        for word in REG_SPACES.split(query):
            if strict:
                words.append(f'"{word}"')
            else:
                words.append(word)
        shell_exec(uri.format(urllib.parse.quote(" ".join(words))))

    return _searcher
