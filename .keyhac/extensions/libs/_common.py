import time
from collections.abc import Callable

CallbackFunc = Callable[[], None]


def delay(msec: int = 50) -> None:
    if 0 < msec:
        time.sleep(msec / 1000)
