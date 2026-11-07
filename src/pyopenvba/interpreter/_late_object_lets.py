"""What a member answers to an object given by Let, measured per member, as live Excel answered it.

Written by scripts/build_late_object_lets.py from tests/fixtures/late_object_lets.json and late_object_lets_2.json.  Do not edit
by hand.

Keyed by binding, then class.member, then the kind of object: novalue (no default member a Let can read), collection, nothing,
cells (one area of several cells) and areas (several areas). Each is the error Excel gives, the member left as it was, or 0 where
the member takes the object's value. A member or kind not here is refused openly: the answer goes by the member, not by its type,
so Workbook.Date1904 and Workbook.CheckCompatibility, both Booleans, answer a Font with 1004 and 13.
"""
from __future__ import annotations

LATE_OBJECT_LETS: dict[str, dict[str, dict[str, int]]] = {
    'typed': {
        'application.statusbar': {'areas': 1004, 'novalue': 1004},
        'font.color': {'cells': 13, 'novalue': 13},
        'font.name': {'cells': 0, 'areas': 1004, 'novalue': 1004},
        'interior.colorindex': {'cells': 13, 'novalue': 13},
        'interior.pattern': {'cells': 13, 'novalue': 13},
        'name.refersto': {'novalue': 438},
        'name.visible': {'cells': 13, 'novalue': 438},
        'range.formula': {'cells': 0, 'novalue': 1004},
        'range.locked': {'cells': 1004, 'novalue': 1004},
        'range.numberformat': {'cells': 0, 'areas': 1004, 'novalue': 1004},
        'range.style': {'cells': 450, 'areas': 450, 'novalue': 450},
        'range.value': {'cells': 0, 'areas': 0, 'novalue': 1004},
        'workbook.checkcompatibility': {'cells': 13, 'novalue': 438},
        'workbook.conflictresolution': {'cells': 13, 'novalue': 438},
        'workbook.date1904': {'cells': 13, 'areas': 0, 'novalue': 438},
        'workbook.updatelinks': {'cells': 13, 'areas': 1004, 'novalue': 438},
        'worksheet.displaypagebreaks': {'cells': 13, 'novalue': 438},
        'worksheet.enableselection': {'cells': 13, 'novalue': 438},
        'worksheet.name': {'cells': 13, 'areas': 0, 'novalue': 438},
        'worksheet.scrollarea': {'cells': 13, 'novalue': 438},
        'worksheet.standardwidth': {'cells': 13, 'areas': 0, 'novalue': 438},
        'worksheet.visible': {'cells': 13, 'areas': 1004, 'novalue': 438},
    },
    'late': {
        'application.statusbar': {'areas': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
        'font.color': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'font.name': {'cells': 0, 'areas': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
        'interior.colorindex': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'interior.pattern': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'name.refersto': {'novalue': 438, 'collection': 450, 'nothing': 5},
        'name.visible': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'range.formula': {'cells': 0, 'novalue': 1004, 'collection': 1004, 'nothing': 1004},
        'range.locked': {'cells': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
        'range.numberformat': {'cells': 0, 'areas': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
        'range.style': {'cells': 450, 'areas': 450, 'novalue': 450, 'collection': 450, 'nothing': 450},
        'range.value': {'cells': 0, 'areas': 0, 'novalue': 1004, 'collection': 1004, 'nothing': 1004},
        'workbook.checkcompatibility': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'workbook.conflictresolution': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'workbook.updatelinks': {'cells': 13, 'areas': 1004, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'worksheet.displaypagebreaks': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'worksheet.enableselection': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'worksheet.name': {'cells': 13, 'areas': 0, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'worksheet.scrollarea': {'cells': 13, 'novalue': 13, 'collection': 13, 'nothing': 13},
        'worksheet.standardwidth': {'areas': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
        'worksheet.visible': {'cells': 1004, 'areas': 1004, 'novalue': 1004, 'collection': 1004, 'nothing': 5},
    },
}
