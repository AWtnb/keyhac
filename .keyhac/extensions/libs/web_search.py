from .text_utils.punctuation import (
    KANGXI_RADICAL_MAPPING,
    RADICAL_MAPPING,
    SEACH_NOISE_MAPPING,
)


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
