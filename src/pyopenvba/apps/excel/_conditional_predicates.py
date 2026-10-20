"""Excel's generated predicate formulas, adapted from pyOfficeEditor.

These are compatibility formulas for attribute-driven rules. Ordinary cellIs
and expression conditions belong to the caller and are never regenerated.
"""
from __future__ import annotations

from pyopenvba._a1 import Area, column_letter
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._xstring import escape
from pyopenvba.exceptions import VBAUnsupportedError

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
TEXT_TYPES = ('containsText', 'notContainsText', 'beginsWith', 'endsWith')
TEXT_OPERATORS = ('containsText', 'notContains', 'beginsWith', 'endsWith')
TIME_PERIODS = ('today', 'yesterday', 'last7Days', 'thisWeek', 'lastWeek', 'lastMonth',
                'tomorrow', 'nextWeek', 'nextMonth', 'thisMonth')
_TEXT_FORMULAS = {
    'containsText': 'NOT(ISERROR(SEARCH("{text}",{anchor})))',
    'notContainsText': 'ISERROR(SEARCH("{text}",{anchor}))',
    'beginsWith': 'LEFT({anchor},LEN("{text}"))="{text}"',
    'endsWith': 'RIGHT({anchor},LEN("{text}"))="{text}"',
}
_PRESENCE_FORMULAS = {
    'containsBlanks': 'LEN(TRIM({anchor}))=0',
    'notContainsBlanks': 'LEN(TRIM({anchor}))>0',
    'containsErrors': 'ISERROR({anchor})',
    'notContainsErrors': 'NOT(ISERROR({anchor}))',
}
_TIME_FORMULAS = {
    'today': 'FLOOR({anchor},1)=TODAY()',
    'yesterday': 'FLOOR({anchor},1)=TODAY()-1',
    'tomorrow': 'FLOOR({anchor},1)=TODAY()+1',
    'last7Days': 'AND(TODAY()-FLOOR({anchor},1)<=6,FLOOR({anchor},1)<=TODAY())',
    'thisWeek': ('AND(TODAY()-ROUNDDOWN({anchor},0)<=WEEKDAY(TODAY())-1,'
                 'ROUNDDOWN({anchor},0)-TODAY()<=7-WEEKDAY(TODAY()))'),
    'lastWeek': ('AND(TODAY()-ROUNDDOWN({anchor},0)>=(WEEKDAY(TODAY())),'
                 'TODAY()-ROUNDDOWN({anchor},0)<(WEEKDAY(TODAY())+7))'),
    'nextWeek': ('AND(ROUNDDOWN({anchor},0)-TODAY()>(7-WEEKDAY(TODAY())),'
                 'ROUNDDOWN({anchor},0)-TODAY()<(15-WEEKDAY(TODAY())))'),
    'thisMonth': 'AND(MONTH({anchor})=MONTH(TODAY()),YEAR({anchor})=YEAR(TODAY()))',
    'lastMonth': ('AND(MONTH({anchor})=MONTH(EDATE(TODAY(),0-1)),'
                  'YEAR({anchor})=YEAR(EDATE(TODAY(),0-1)))'),
    'nextMonth': ('AND(MONTH({anchor})=MONTH(EDATE(TODAY(),0+1)),'
                  'YEAR({anchor})=YEAR(EDATE(TODAY(),0+1)))'),
}
PREDICATE_TYPES = frozenset((*TEXT_TYPES, *_PRESENCE_FORMULAS, 'timePeriod'))


def formula(kind: str, anchor: Area, *, text: str | None = None, time_period: str | None = None) -> str | None:
    reference = column_letter(anchor.left) + str(anchor.top)
    template = _TEXT_FORMULAS.get(kind)
    if template is not None:
        return None if text is None else template.format(anchor=reference, text=text.replace('"', '""'))
    template = _PRESENCE_FORMULAS.get(kind)
    if kind == 'timePeriod' and time_period is not None:
        template = _TIME_FORMULAS.get(time_period)
    return None if template is None else template.format(anchor=reference)


def formula_for_rule(root: XmlElement, anchor: Area) -> str | None:
    text = root.get('text')
    return formula(root.get('type', ''), anchor, text=text,
                   time_period=root.get('timePeriod'))


def saved_text(text: str) -> str:
    """The plain XML attribute drops CR; its formula keeps CR escaped.

    Native Excel can create an in-memory NUL predicate but cannot reopen its
    own saved workbook. Refuse persistence instead of emitting that corruption.
    """
    if '\x00' in text:
        raise VBAUnsupportedError('Saving conditional text containing NUL is not supported by native Excel')
    text = text.replace('\r', '').encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'ignore')
    if any((ord(character) < 32 and character not in '\t\n') or character in '\ufffe\uffff' for character in text):
        raise VBAUnsupportedError('Saving this non-XML conditional text character is not implemented')
    return text


def regenerate(root: XmlElement, anchor: Area) -> None:
    expression = formula_for_rule(root, anchor)
    if expression is None:
        return
    original = root.findall(NS + 'formula')
    if original:
        original[0].text = escape(expression)
        for extra in original[1:]:
            root.remove(extra)
    else:
        child = XmlElement(NS + 'formula')
        child.text = escape(expression)
        root.insert(0, child)
