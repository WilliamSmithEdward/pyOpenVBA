"""Generate src/pyopenvba/interpreter/_late_lets.py: what a call bound late answers where it differs from typed.

Typed, VBA turns a value into the parameter's type before the call, so that Null is 94, text that is no number 13 and True -1. Late,
the object takes the Variant itself. No rule in the type library says which answer a member gives, so the answers are read from
the fixtures that measured both, per class.member and kind of value.

    python scripts/build_late_lets.py
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests/fixtures'
OUT = ROOT / 'src/pyopenvba/interpreter/_late_lets.py'
METHOD_FIXTURES = ('method_audit', 'range_methods_clipboard', 'range_methods_moves', 'range_methods_order', 'range_methods_search')


def kind_of(value: str) -> str | None:
    if value == 'Null':
        return 'null'
    if value in ('True', 'False'):
        return 'bool'
    return 'text' if value.startswith('"') else None


def number_of(answer: str) -> int | None:
    found = re.match(r'[eE](-?\d+)', answer)
    return int(found.group(1)) if found else None


def note(table: dict[str, dict[str, int]], key: str, kind: str | None, early: int | None, late: int | None) -> None:
    if kind is not None and early is not None and late is not None and early != late:
        table.setdefault(key.lower(), {})[kind] = late


def lets() -> dict[str, dict[str, int]]:
    table: dict[str, dict[str, int]] = {}
    for name, row in json.loads((FIXTURES / 'settable_audit.json').read_text(encoding='utf-8'))['rows'].items():
        count = len(row['values'])
        for value, late, typed in zip(row['values'], row['answers'][:count], row['answers'][count:]):
            if late.startswith('failed') or typed.startswith('failed') or '|' not in late or '|' not in typed:
                continue
            early = number_of(typed.split('|')[1])
            found = number_of(late.split('|')[1])
            if early == 5 and found is not None:
                table.setdefault(name.lower(), {})['five'] = found
            else:
                note(table, name, kind_of(value), early, found)
    for fixture in ('member_audit', 'format_properties'):
        record = json.loads((FIXTURES / f'{fixture}.json').read_text(encoding='utf-8'))
        for number, early in record['early'].items():
            case = record['cases'][int(number)]
            note(table, f'{case[0]}.{case[2]}', kind_of(case[3]), number_of(early), number_of(case[4]))
    record = json.loads((FIXTURES / 'range_more.json').read_text(encoding='utf-8'))
    for number, early in record['early'].items():
        case = record['cases'][int(number)]
        note(table, f'Range.{case[0]}', kind_of(case[3]), number_of(early), number_of(case[4]))
    return table


def calls() -> dict[str, dict[str, int]]:
    table: dict[str, dict[str, int]] = {}
    for fixture in METHOD_FIXTURES:
        record = json.loads((FIXTURES / f'{fixture}.json').read_text(encoding='utf-8'))
        for number, early in record['early'].items():
            case = record['cases'][int(number)]
            kind = 'null' if re.search(r'\bNull\b', case[2]) else 'bool' if re.search(r'\b(True|False)\b', case[2]) else 'text' if '"' in case[2] else None
            # AutoFilter's answer goes by which argument it is, which a kind of value does not say.
            if case[1] != 'AutoFilter':
                note(table, f'{case[0]}.{case[1]}', kind, number_of(early), number_of(case[3]))
    return table


def indexes() -> dict[str, int]:
    """The members of tests/fixtures/late_indexes.json that refuse an index. Any other answer is what they answered being indexed:
    a 9 is an empty collection's Item saying so, and a 438 the Let of an object with no default, as Borders(1) reads."""
    record = json.loads((FIXTURES / 'late_indexes.json').read_text(encoding='utf-8'))
    return {f'{name.lower()}.{member}': number for name, members in record['classes'].items() for member, number in members.items()
            if number in (450, 451)}


def main() -> None:
    lines = ['"""What a call bound late answers where it differs from typed, measured per class.member.', '',
             'Written by scripts/build_late_lets.py.  Do not edit by hand.', '',
             'Keyed by kind of value: null, text and bool where typed VBA answers 94, 13 and 1004 before the call, and five for a value the',
             'member refuses with 5. A 0 means the member takes the value.', '"""', 'from __future__ import annotations', '', 'LATE_LETS: dict[str, dict[str, int]] = {']
    built = lets()
    lines += [f'    {key!r}: {built[key]!r},' for key in sorted(built)]
    lines += ['}', '', 'LATE_CALLS: dict[str, dict[str, int]] = {']
    made = calls()
    lines += [f'    {key!r}: {made[key]!r},' for key in sorted(made)]
    lines += ['}', '', '#: A property that takes nothing, given an index (tests/fixtures/late_indexes.json).', 'LATE_INDEXES: dict[str, int] = {']
    numbered = indexes()
    lines += [f'    {key!r}: {numbered[key]!r},' for key in sorted(numbered)]
    lines += ['}', '', '#: Read through an Object these are Doubles, where the library declares them Long (tests/fixtures/member_audit.json,',
              '#: tests/fixtures/excel_everyday_16.json, tests/fixtures/old_collection_members.json, tests/fixtures/typed_answers/late_types_ControlFormat.json).',
              "LATE_DOUBLES = frozenset({'application.maxiterations', 'application.sheetsinnewworkbook', 'workbook.fileformat',",
              "                          'dropdown.listcount', 'dropdown.listindex', 'dropdown.value', 'dropdown.dropdownlines',",
              "                          'listbox.listcount', 'listbox.listindex', 'listbox.value', 'controlformat.dropdownlines',",
              "                          'controlformat.listcount', 'controlformat.listindex', 'controlformat.max', 'controlformat.min',",
              "                          'controlformat.smallchange'})", '']
    OUT.write_text('\n'.join(lines), encoding='utf-8', newline='\n')
    print(len(built), 'properties,', len(made), 'methods')


if __name__ == '__main__':
    main()
