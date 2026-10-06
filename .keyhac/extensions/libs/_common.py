import fnmatch
import time
from collections.abc import Callable

type CallbackFunc = Callable[[], None]


def delay(msec: int = 50) -> None:
    if 0 < msec:
        time.sleep(msec / 1000)


def remove_extension(s: str) -> str:
    period = "."
    parts = s.split(period)
    ext = parts.pop()
    if not parts:
        return ext
    return period.join(parts)


def is_same_app(a: str, b: str) -> bool:
    return fnmatch.fnmatch(remove_extension(a), remove_extension(b))
