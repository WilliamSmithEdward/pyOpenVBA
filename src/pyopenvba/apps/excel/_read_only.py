"""The errors from assigning to what Excel holds read-only, measured member by member on late-bound objects
(tests/fixtures/read_only_lets.json). Most answer 450; these do not, and nothing in the type library says which."""
from __future__ import annotations

_OTHER_LET_ERRORS: dict[int, dict[str, tuple[str, ...]]] = {
    1004: {
        'application': ('build', 'caller', 'canplaysounds', 'canrecordsounds', 'customlistcount', 'international',
                        'mailsystem', 'mathcoprocessoravailable', 'mouseavailable', 'operatingsystem', 'path',
                        'pathseparator', 'previousselections', 'recordrelative', 'version'),
        'chart': ('codename', 'index', 'protectcontents', 'protectdrawingobjects', 'protectionmode'),
        'filter': ('count', 'criteria1', 'criteria2'),
        'formatcondition': ('formula2',),
        'listobject': ('sharepointurl',),
        'pivotcache': ('isconnected', 'querytype', 'sourcedatafile'),
        'pivotcell': ('customsubtotalfunction',),
        'pivotfield': ('datatype', 'displayascaption', 'grouplevel', 'ismemberproperty', 'sourcename', 'totallevels'),
        'pivotfields': ('count',),
        'pivotitem': ('parentshowdetail', 'sourcename'),
        'pivotitems': ('count',),
        'pivottable': ('mdx',),
        'pivottables': ('count',),
        'range': ('hasarray', 'hasformula', 'height', 'left', 'listheaderrows', 'locationintable', 'prefixcharacter',
                  'summary', 'text', 'top', 'width'),
        'scenario': ('index', 'values'),
        'scenarios': ('count',),
        'shape': ('formcontroltype',),
        'style': ('namelocal',),
        'window': ('index', 'type', 'windownumber'),
        'workbook': ('createbackup', 'fileformat', 'fullname', 'fullnameurlencoded', 'haspassword', 'multiuserediting',
                     'path', 'readonly', 'revisionnumber', 'writereserved'),
        'worksheet': ('consolidationfunction', 'consolidationoptions', 'consolidationsources', 'protectcontents',
                      'protectdrawingobjects', 'protectscenarios', 'standardheight', 'type'),
    },
    438: {
        'pivotfield': ('hiddenitems', 'parentitems', 'visibleitems'),
        'pivotitem': ('childitems',),
        'range': ('address', 'addresslocal'),
    },
    5: {'pivotfield': ('sourcecaption',)},
}
_LET_ERRORS = {(owner, name): number for number, owners in _OTHER_LET_ERRORS.items()
               for owner, names in owners.items() for name in names}


def let_error(owner: str, name: str) -> int:
    """The error from assigning a value to a read-only member of the type called owner."""
    return _LET_ERRORS.get((owner.lower(), name.lower()), 450)


def set_parent_error(owner: str, parent: str) -> int:
    """The error from `Set o.Parent = x`, which follows what Parent answers: 450 on the three types below, 13 under a
    pivot field and for a pivot table's fields, and 438 on the rest, pivot tables' other children included."""
    if parent in ('Application', 'ListObject', 'Style'):
        return 450
    if parent == 'PivotField' or (parent == 'PivotTable' and owner in ('PivotField', 'PivotFields')):
        return 13
    return 438
