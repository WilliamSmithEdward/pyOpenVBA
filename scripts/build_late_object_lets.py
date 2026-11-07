"""Write src/pyopenvba/interpreter/_late_object_lets.py from what live Excel answered to an object given by Let.

Reads tests/fixtures/late_object_lets.json and late_object_lets_2.json. For each member, and each kind of object a member answers by its own
rule (one with no value, Nothing, one area of several cells, several areas), it keeps Excel's error where the set changed nothing, and leaves
the kind out where the set changed the member or took the object, so the model refuses that case openly. Typed and late answers are kept
apart, as a typed member of a value type gets the object's default value from VBA instead.

    python scripts/build_late_object_lets.py
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'src/pyopenvba/interpreter/_late_object_lets.py'
#: The kind each probe object is, for a member that receives the object.
KINDS = {'ws.Range("A1").Font': 'novalue', 'ws.Range("A1").Interior': 'novalue', 'ws': 'novalue', 'New Collection': 'collection',
         'Nothing': 'nothing', 'ws.Range("A1:C1")': 'cells', 'ws.Range("A1:B1")': 'cells', 'ws.Range("A1,C1")': 'areas'}


def build() -> dict[str, dict[str, dict[str, int]]]:
    table: dict[str, dict[str, dict[str, int]]] = {'typed': {}, 'late': {}}
    refused: dict[str, set[tuple[str, str]]] = {'typed': set(), 'late': set()}
    for name in ('late_object_lets', 'late_object_lets_2'):
        record = json.loads((ROOT / f'tests/fixtures/{name}.json').read_text(encoding='utf-8'))
        for binding in ('typed', 'late'):
            answers = record['excel'][binding]
            # What the member reads when the set changes nothing: the read after the first refused set of that member.
            unchanged: dict[str, set[str]] = {}
            for case, answer in zip(record['cases'], answers):
                member = f'{case[0].lower()}.{case[2].lower()}'
                if answer != 'does not compile' and not answer.startswith('e0'):
                    unchanged.setdefault(member, set()).add(answer.split('/', 1)[1])
            before = record['predicted' if 'predicted' in record else 'model'][binding]
            for case, answer, model in zip(record['cases'], answers, before):
                kind = KINDS.get(case[-1])
                member = f'{case[0].lower()}.{case[2].lower()}'
                if kind is None or answer == 'does not compile':
                    continue
                number, read = answer.split('/', 1)
                code = int(number.removeprefix('e'))
                # Kept where the set refused and left the member as other refused sets of it left it, or, as 0, where Excel took the
                # object's value just as the model's value-taking path did.
                if not code and model == answer:
                    code = 0
                if (not code and model == answer) or (code and len(unchanged.get(member, ())) == 1 and read in unchanged[member]):
                    entry = table[binding].setdefault(member, {})
                    if entry.get(kind, code) != code:
                        refused[binding].add((member, kind))
                    entry[kind] = code
                else:
                    refused[binding].add((member, kind))
    for binding, pairs in refused.items():
        for member, kind in pairs:
            table[binding].get(member, {}).pop(kind, None)
    return {binding: {member: kinds for member, kinds in sorted(members.items()) if kinds} for binding, members in table.items()}


def main() -> None:
    table = build()
    lines = ['"""What a member answers to an object given by Let, measured per member, as live Excel answered it.', '',
             'Written by scripts/build_late_object_lets.py from tests/fixtures/late_object_lets.json and late_object_lets_2.json.  Do not edit',
             'by hand.', '',
             'Keyed by binding, then class.member, then the kind of object: novalue (no default member a Let can read), collection, nothing,',
             'cells (one area of several cells) and areas (several areas). Each is the error Excel gives, the member left as it was, or 0 where',
             'the member takes the object\'s value. A member or kind not here is refused openly: the answer goes by the member, not by its type,',
             'so Workbook.Date1904 and Workbook.CheckCompatibility, both Booleans, answer a Font with 1004 and 13.', '"""',
             'from __future__ import annotations', '',
             'LATE_OBJECT_LETS: dict[str, dict[str, dict[str, int]]] = {']
    for binding, members in table.items():
        lines.append(f'    {binding!r}: {{')
        lines += [f'        {member!r}: {kinds!r},' for member, kinds in members.items()]
        lines.append('    },')
    lines.append('}')
    OUT.write_text('\n'.join(lines) + '\n', encoding='utf-8', newline='\n')
    print(sum(len(members) for members in table.values()), 'members written')


if __name__ == '__main__':
    main()
