import re
from collections.abc import Callable

from keyhac import *  # type: ignore

from .punctuation import KANGXI_RADICAL_MAPPING, RADICAL_MAPPING
from .str_tools import (
    colon_to_doubledash,
    decode_url,
    encode_url,
    fix_dumb_quotation,
    fix_paren_inside_bracket,
    format_nested_bracket,
    format_nested_paren,
    insert_blank_line,
    invoke_comment_remover,
    mdtable_from_tsv,
    remove_whitespace,
    skip_blank_line,
    split_postalcode,
    swap_abbreviation,
    swap_tabs,
    to_double_bracket,
    to_full_brackets,
    to_full_letter,
    to_full_symbol,
    to_half_brackets,
    to_half_letter,
    to_half_symbol,
    to_list,
    to_single_bracket,
    trim_honorific,
    trim_space_on_line_head,
)

CLIPBOARD_FORMATTER_MAPPING = {}


def set_formatter(binding: dict) -> None:
    for menu, func in binding.items():
        CLIPBOARD_FORMATTER_MAPPING[menu] = func


set_formatter(
    {
        "to codeblock": lambda c: f"```\n{c}\n```\n",
        "swap tabs": swap_tabs,
        "trim space on line head": trim_space_on_line_head,
        "to lowercase": lambda c: c.lower(),
        "to uppercase": lambda c: c.upper(),
        "to slack feed subscribe": lambda c: f"/feed subscribe {c}",
        "to slack feed remove": lambda c: f"/feed remove {c}",
        "to list": to_list,
        "swap abbreviation around colon": swap_abbreviation,
        "colon to double-dash": colon_to_doubledash,
        "insert blank line": insert_blank_line,
        "remove blank line": skip_blank_line,
        "fix dumb quotation": fix_dumb_quotation,
        "fix KANGXI RADICALS": lambda s: s.transrate(
            str.maketrans(KANGXI_RADICAL_MAPPING | RADICAL_MAPPING)
        ),
        "fix paren inside bracket": fix_paren_inside_bracket,
        "to double bracket": to_double_bracket,
        "to single bracket": to_single_bracket,
        "TSV to markdown table": mdtable_from_tsv,
        "split postalcode and address": split_postalcode,
        "decode url": decode_url,
        "encode url": encode_url,
        "to halfwidth": lambda s: to_half_letter(s, False),
        "to halfwidth (including symbols)": lambda s: to_half_letter(s, True),
        "to halfwidth symbols": to_half_symbol,
        "to halfwidth bracktets": to_half_brackets,
        "to fullwidth": lambda s: to_full_letter(s, False),
        "to fullwidth (including symbols)": lambda s: to_full_letter(s, True),
        "to fullwidth symbols": to_full_symbol,
        "to fullwidth bracktets": to_full_brackets,
        "trim honorific": trim_honorific,
        "fix nested paren": format_nested_paren,
        "fix nested bracket": format_nested_bracket,
        "remove whitespaces": remove_whitespace,
        "remove javascript comment line": invoke_comment_remover("// "),
        "remove python comment line": invoke_comment_remover("# "),
    }
)


def invoke_replacer(search: str, replace_to: str) -> Callable[[str], str]:
    reg = re.compile(search)

    def _replacer(s: str) -> str:
        return reg.sub(replace_to, s)

    return _replacer


def set_replacer(binding: dict) -> None:
    for menu, args in binding.items():
        CLIPBOARD_FORMATTER_MAPPING[menu] = invoke_replacer(*args)


set_replacer(
    {
        "backslash to slash": (r"\\", "/"),
        "escape backslash": (r"\\", r"\\\\"),
        "escape double-quotation": (r"\"", r'\\"'),
        "remove double-quotation": (r'"', ""),
        "remove single-quotation": (r"'", ""),
        "remove linebreak": (r"\r?\n", ""),
        "to sigle line": (r"\r?\n", ""),
        "remove whitespaces (including linebreak)": (r"\s", ""),
        "remove non-digit-char": (r"[^\d]", ""),
        "remove quotations": (r"[\u0022\u0027]", ""),
        "remove inside paren": (r"[（\(].+?[）\)]", ""),
        "fix msword-bullet": (
            r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009",
            "\u30fb",
        ),
        "remove msword-bullet": (
            r"[\uF06C\uF0D8\uF0B2\uF09F\u25E6\uF0A7\uF06C]\u0009",
            "",
        ),
        "to curly-comma (\uff0c)": (r"\u3001", "\uff0c"),
        "to japanese-comma (\u3001)": (r"\uff0c", "\u3001"),
        "shorten amazon url": (
            r"^.+amazon\.co\.jp/.+dp/(.{10}).*",
            r"https://www.amazon.jp/dp/\1",
        ),
    }
)


def invoke_line_jointer(sep: str) -> Callable[[str], str]:
    def _jointer(s: str) -> str:
        return sep.join(s.splitlines())

    return _jointer


def set_line_jointer(binding: dict) -> None:
    for name, sep in binding.items():
        menu = f"Join lines with {name}"
        CLIPBOARD_FORMATTER_MAPPING[menu] = invoke_line_jointer(sep)


set_line_jointer(
    {
        "Half-Comma": ",",
        "Dot": "・",
        "Tab": "\t",
        "Slash": "／",
        "Pipe": "|",
    }
)
