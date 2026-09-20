"""Things in the source that only work on one platform.

pyOpenVBA runs wherever Python does; the Office files it reads were
written by Windows, but nothing here may need Windows to read them.
This is the check for the ways that quietly stops being true, each one
found the hard way in CI rather than here.
"""

from __future__ import annotations

import re
from pathlib import Path

SOURCE = Path(__file__).resolve().parent.parent / "src" / "pyopenvba"

#: strftime's way of asking for a number without its leading zero is
#: platform-specific: %#m is Windows, %-m is glibc and BSD, and neither
#: is the other's.  A date written with one is right on the machine it
#: was written on and wrong everywhere else, which is exactly the bug
#: that gets past a Windows-only run.
_PADDING = re.compile(r"%[#-][aAbBcdHIjmMpSUwWxXyYzZ]")

#: Where a date format can be written: a strftime call, or the format
#: spec of an f-string such as f"{when:%H:%M}".
_FORMATTING = re.compile(r"strftime|:%")


def _sources() -> list[Path]:
    # The packaged file templates are base85 blobs, not code.
    return [path for path in sorted(SOURCE.rglob("*.py")) if "_templates" not in path.parts]


def test_no_platform_specific_strftime() -> None:
    found: list[str] = []
    for path in _sources():
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if _PADDING.search(line) and _FORMATTING.search(line):
                found.append(f"{path.relative_to(SOURCE)}:{number}: {line.strip()}")
    assert not found, "strftime padding flags are platform-specific:\n" + "\n".join(found)
