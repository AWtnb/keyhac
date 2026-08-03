import re
import urllib.parse
from collections.abc import Callable

from .char_width import to_half_letter


def remove_whitespace(s: str) -> str:
    return s.strip().translate(
        str.maketrans(
            "",
            "",
            "\u0009\u0020\u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u200c\u200d\u200e\u200f\u202f\u205f\u3000\ufeff",
        )
    )


def simple_quote(s: str) -> str:
    lines = s.strip().splitlines()
    return "\n".join([">" + line for line in lines])


def as_single_quoted_line(s: str) -> str:
    lines = s.strip().splitlines()
    return ">" + "".join([line.strip() for line in lines])


def invoke_comment_remover(symbol: str) -> Callable[[str], str]:
    def _remover(s: str) -> str:
        return "\n".join(
            [line for line in s.splitlines() if not line.strip().startswith(symbol)]
        )

    return _remover


class NestedCircumfix:
    def __init__(self, prime_pair: tuple, secondary_pair: tuple):
        self.pairs = [prime_pair, secondary_pair]

    def fix(self, s: str) -> str:
        stack = []
        result = list(s)
        openChar, closeChar = self.pairs[0]
        for i, char in enumerate(s):
            if char == openChar:
                stack.append(i)
            else:
                if char == closeChar and stack:
                    start = stack.pop()
                    depth = len(stack)
                    left, right = self.pairs[depth % 2]
                    result[start] = left
                    result[i] = right

        return "".join(result)


def swap_abbreviation(s: str) -> str:
    ss = re.split(r"[:：]\s*", s)
    if len(ss) == 2:
        return ss[1] + "：" + ss[0]
    return ""


def colon_to_doubledash(s: str) -> str:
    return re.sub(r"[:：]\s*", "\u2015\u2015", s)


def skip_blank_line(s: str) -> str:
    lines = s.strip().splitlines()
    return "\n".join([line for line in lines if line.strip()])


def insert_blank_line(s: str) -> str:
    lines = []
    for line in s.strip().splitlines():
        lines.append(line.strip())
        lines.append("")
    return "\n".join(lines)


def to_double_bracket(s: str) -> str:
    reg = re.compile(r"[\u300c\u300d]")

    def _replacer(mo: re.Match) -> str:
        if mo.group(0) == "\u300c":
            return "\u300e"
        return "\u300f"

    return reg.sub(_replacer, s)


def to_single_bracket(s: str) -> str:
    reg = re.compile(r"[\u300e\u300f]")

    def _replacer(mo: re.Match) -> str:
        if mo.group(0) == "\u300e":
            return "\u300c"
        return "\u300d"

    return reg.sub(_replacer, s)


def to_list(s: str) -> str:
    lines = s.splitlines()
    return "\n".join(["- " + line for line in lines])


def split_postalcode(s: str) -> str:
    lines = s.splitlines()
    if 1 < len(lines):
        reg = re.compile(r"(\d{3}).(\d{4})[ 　]*(.+$)")
    else:
        reg = re.compile(r"(\d{3}).(\d{4})[\s]*(.+$)")
    ss = []
    for line in lines:
        hankaku = to_half_letter(line.strip().strip("\u3012"), True)
        m = reg.match(hankaku)
        if m:
            ss.append(f"{m.group(1)}-{m.group(2)}\t{m.group(3)}")
        else:
            ss.append(line)
    return "\n".join(ss)


def fix_paren_inside_bracket(s: str) -> str:
    reg = re.compile(r"(\(.+?\)|（.+?）)」")

    def _replacer(mo: re.Match) -> str:
        return "」" + mo.group(1)

    return reg.sub(_replacer, s)


def fix_dumb_quotation(s: str) -> str:
    reg = re.compile(r"\"([^\"]+?)\"|'([^']+?)'")

    def _replacer(mo: re.Match) -> str:
        if str(mo.group(0)).startswith('"'):
            return f"\u201c{mo.group(1)}\u201d"
        return f"\u2018{mo.group(1)}\u2019"

    return reg.sub(_replacer, s)


def decode_url(s: str) -> str:
    return urllib.parse.unquote(s)


def encode_url(s: str) -> str:
    return urllib.parse.quote(s)


def trim_honorific(s: str) -> str:
    reg = re.compile(r"先生$|様$|(先生|様)(?=[、。：；（）［］・！？\s])")
    return reg.sub("", s)


def trim_space_on_line_head(s: str) -> str:
    return "\n".join([line.lstrip() for line in s.splitlines()])


def format_nested_paren(s: str) -> str:
    return NestedCircumfix(("（", "）"), ("〔", "〕")).fix(s)


def format_nested_bracket(s: str) -> str:
    return NestedCircumfix(("「", "」"), ("『", "』")).fix(s)


def swap_tabs(s: str) -> str:
    lines = s.splitlines()
    if len(lines) < 1:
        return s
    swapped = []
    for line in lines:
        ss = line.split("\t")
        ss.insert(0, ss.pop())
        swapped.append("\t".join(ss))
    return "\n".join(swapped)


def mdtable_from_tsv(s: str) -> str:
    delim = "\t"

    def _split(s: str) -> list[str]:
        return s.split(delim)

    def _join(ss: list) -> str:
        pipe = "|"
        return pipe + pipe.join(ss) + pipe

    lines = s.splitlines()
    header = _join(_split(lines[0]))
    sep = _join([":---:" for _ in lines[0].split(delim)])
    table = [
        header,
        sep,
    ]
    for line in lines[1:]:
        table.append(_join(_split(line)))
    return "\n".join(table)
