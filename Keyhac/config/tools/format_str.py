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


def as_single_line(s: str) -> str:
    lines = s.strip().splitlines()
    return ">" + "".join([line.strip() for line in lines])


def skip_blank_line(s: str) -> str:
    lines = []
    for line in s.strip().splitlines():
        if 0 < len(line.strip()):
            lines.append(">" + line)
        else:
            lines.append("")
    return "\n".join(lines)
