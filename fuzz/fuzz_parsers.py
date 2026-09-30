"""Coverage-guided fuzzing of the parsers that read untrusted files.

Each target feeds Atheris-generated input to one parser. A parser must
either succeed or raise one of the exceptions tests/fuzz_corpus/README.md
lists for it; anything else, a crash or a hang, is a finding. Atheris
writes the input that caused it to a crash-* file.

    python fuzz/fuzz_parsers.py <target> [libFuzzer options] [corpus dirs]

    python fuzz/fuzz_parsers.py decompress -max_total_time=60 tests/fuzz_corpus/decompress

Targets: cfb, decompress, dir, project, projectwm, mashup, mformula.
The .github/workflows/fuzz.yml workflow runs each one, seeded from
tests/fuzz_corpus/<target> where that exists. A finding becomes a
regression seed in that directory, which tests/test_gates.py replays on
every CI run.
"""

import struct
import sys

import atheris

with atheris.instrument_imports():
    from pyopenvba.cfb import CFB
    from pyopenvba.exceptions import CFBError, PowerQueryError, PyOpenVBAError, VBAProjectError
    from pyopenvba.mlang._parse import MSyntaxError, parse
    from pyopenvba.powerquery._mashup import Mashup
    from pyopenvba.vba import (
        _parse_dir_stream,
        decompress,
        parse_project_stream,
        parse_projectwm,
        parse_vba_project,
    )

# A cap on decompressed output, so a small input that claims a huge size is
# refused rather than exhausting memory.
MAX_BYTES = 1 << 22


def fuzz_cfb(data):
    try:
        parse_vba_project(CFB.from_bytes(data))
    except (CFBError, PyOpenVBAError, UnicodeDecodeError):
        pass


def fuzz_decompress(data):
    try:
        decompress(data, max_bytes=MAX_BYTES)
    except VBAProjectError:
        pass


def fuzz_dir(data):
    try:
        _parse_dir_stream(data)
    except (VBAProjectError, UnicodeDecodeError, IndexError, struct.error):
        pass


def fuzz_project(data):
    try:
        parse_project_stream(data)
    except (VBAProjectError, UnicodeDecodeError):
        pass


def fuzz_projectwm(data):
    try:
        parse_projectwm(data)
    except (VBAProjectError, UnicodeDecodeError):
        pass


def fuzz_mashup(data):
    try:
        Mashup.parse(data)
    except PowerQueryError:
        pass


def fuzz_mformula(data):
    source = atheris.FuzzedDataProvider(data).ConsumeUnicodeNoSurrogates(len(data))
    try:
        parse(source)
    except MSyntaxError:
        pass


TARGETS = {
    "cfb": fuzz_cfb,
    "decompress": fuzz_decompress,
    "dir": fuzz_dir,
    "project": fuzz_project,
    "projectwm": fuzz_projectwm,
    "mashup": fuzz_mashup,
    "mformula": fuzz_mformula,
}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_parsers.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    target = TARGETS[sys.argv[1]]
    atheris.Setup([sys.argv[0], *sys.argv[2:]], target)
    atheris.Fuzz()


if __name__ == "__main__":
    main()
