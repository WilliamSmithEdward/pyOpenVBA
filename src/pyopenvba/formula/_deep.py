"""Room for the deepest formula Excel takes.

Excel takes a formula 8192 characters long: 4096 terms added one to the
next, 256 brackets one inside another, 65 functions, a thousand signs
before a number or 8190 percent signs after one. Reading such a formula,
spelling it and working it out go down its tree one call inside
another, far deeper than Python's stack allows by default.
:func:`deep` does the work as usual, and if it runs out of stack, again
on a thread of its own with the stack and the recursion limit the
deepest formula needs. The work has to be one that can be done again:
reading, spelling or working out a formula, not writing it somewhere.
"""

from __future__ import annotations

import sys
import threading
from collections.abc import Callable
from typing import TypeVar

T = TypeVar("T")

#: The stack of the thread the deepest formula is worked on, and Python's recursion limit there.
_STACK = 128 * 1024 * 1024
_LIMIT = 200_000
_THREAD = "pyopenvba-deep-formula"


def deep(work: Callable[[], T]) -> T:
    """What ``work`` gives, done on a deep enough stack; whatever it raises, raised here."""
    try:
        return work()
    except RecursionError:
        if threading.current_thread().name == _THREAD:
            # Already as deep as a formula goes: the work recurses without end.
            raise
    done: list[T] = []
    failed: list[BaseException] = []

    def run() -> None:
        limit = sys.getrecursionlimit()
        sys.setrecursionlimit(max(limit, _LIMIT))
        try:
            done.append(work())
        except BaseException as failure:  # noqa: BLE001 - raised again in the thread that asked
            failed.append(failure)
        finally:
            sys.setrecursionlimit(limit)

    size = threading.stack_size()
    threading.stack_size(_STACK)
    try:
        thread = threading.Thread(target=run, name=_THREAD)
        thread.start()
    finally:
        threading.stack_size(size)
    thread.join()
    if failed:
        raise failed[0]
    return done[0]


__all__ = ["deep"]
