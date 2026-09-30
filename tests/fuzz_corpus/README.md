# Persistent Fuzz Corpus

This directory holds seed inputs for the parsers in `pyopenvba`. Each
subdirectory targets one parser and is walked by the harness in
[`tests/test_gates.py::TestGate23_Fuzz::test_persistent_fuzz_corpus`](../test_gates.py).

For every input file, the harness asserts that the parser **either
succeeds or raises one of the documented exception types** (`CFBError`,
`VBAProjectError`, `PyOpenVBAError`, `UnicodeDecodeError`,
`struct.error`, `IndexError`). Any other exception, or a silent hang /
crash, is treated as a fuzz regression.

## Layout

| Subdir | Target | Acceptable exceptions |
|--------|--------|-----------------------|
| `cfb/` | `CFB.from_bytes()` | `CFBError`, `PyOpenVBAError` |
| `decompress/` | `pyopenvba.vba.decompress()` | `VBAProjectError` |
| `dir/` | `pyopenvba.vba._parse_dir_stream()` | `VBAProjectError`, `UnicodeDecodeError`, `IndexError`, `struct.error` |
| `project/` | `pyopenvba.vba.parse_project_stream()` | `VBAProjectError`, `UnicodeDecodeError` |
| `projectwm/` | `pyopenvba.vba.parse_projectwm()` | `VBAProjectError`, `UnicodeDecodeError` |
| `mashup/` | `pyopenvba.powerquery._mashup.Mashup.parse()` | `PowerQueryError` |
| `mformula/` | `pyopenvba.mlang._parse.parse()` | `MSyntaxError` |

Each input file is **raw bytes** (extension is informational only;
`.bin` by convention). Filenames should describe the seed, e.g.
`empty.bin`, `truncated_header.bin`, `regression_2026_05_xxxx.bin`.

## Coverage-guided fuzzing

[`fuzz/fuzz_parsers.py`](../../fuzz/fuzz_parsers.py) runs each target under
Atheris, seeded from its directory here, with the same acceptable
exceptions. The Fuzz workflow runs every target daily and on each change to
the parsers. To run one yourself on Linux with Python 3.12 or later:

```
python -m pip install --require-hashes -r .github/requirements/fuzz.txt
python -m pip install --no-deps --no-build-isolation -e .
python fuzz/fuzz_parsers.py decompress -max_total_time=60 tests/fuzz_corpus/decompress
```

A finding is written to a `crash-*` file. Minimise it, add it here as a
regression seed, and fix the parser.

## Adding a new seed

When a fuzz run discovers a new failure mode, drop the offending bytes
into the appropriate subdirectory and commit. The harness will pick it
up on next run. Keep individual files small (< 64 KB) and prefer
**minimised** reproducers.

## Regenerating the initial seed set

```
python scripts/seed_fuzz_corpus.py
```

That script is **idempotent and additive** — it never deletes existing
files, only creates any missing initial seeds. Hand-added regression
seeds are safe.
