"""Which members end Copy and Cut. Excel's list follows no rule found, so it stands here as it was measured, object by
object (tests/fixtures/excel_everyday_7.json and excel_everyday_8.json). Font's Bold ends them and its Color does not;
Interior never does; Calculate ends them and CalculateFull does not. Whatever is not here leaves them.
"""
from __future__ import annotations

_RANGE = frozenset({
    # Methods.
    "advancedfilter", "autofill", "autofilter", "borderaround", "clear", "clearcontents", "clearformats",
    "clearhyperlinks", "clearnotes", "clearoutline", "consolidate", "dataseries", "filldown", "fillleft",
    "fillright", "fillup", "group", "justify", "merge", "removeduplicates", "sort", "subtotal", "texttocolumns",
    "ungroup", "unmerge",
    # Properties.
    "addindent", "formulaarray", "formulahidden", "horizontalalignment", "indentlevel", "locked", "numberformat",
    "numberformatlocal", "orientation", "shrinktofit", "style", "verticalalignment", "wraptext",
})
_FONT = frozenset({"bold", "italic", "name", "size", "underline"})
_LINE = frozenset({"linestyle"})
_SHEETS = frozenset({"add"})

ENDING: dict[str, frozenset[str]] = {
    "Application": frozenset({"calculate", "calculation", "iteration"}),
    "Border": _LINE,
    "Borders": _LINE,
    "Font": _FONT,
    "ListColumns": frozenset({"add"}),
    "ListObject": frozenset({"showtotals"}),
    "ListRow": frozenset({"delete"}),
    "ListRows": frozenset({"add"}),
    "PageSetup": frozenset({"centerheader", "leftmargin", "orientation"}),
    "Range": _RANGE,
    "Sheets": _SHEETS,
    "Styles": frozenset({"add"}),
    "Workbook": frozenset({"protect"}),
    "Worksheet": frozenset({"calculate", "copy", "delete", "showalldata"}),
    "Worksheets": _SHEETS,
}
