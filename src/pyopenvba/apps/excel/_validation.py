"""Data validation: Range.Validation, and the dataValidations element a sheet keeps its rules in.

Measured in live Excel (scripts/measure_validation.py, tests/fixtures/validation/):

* A rule covers a set of blocks, its sqref, and is added with
  Validation.Add over cells that have none; over any that has one it is
  error 1004. A rule with the same settings as one the sheet has joins it,
  its blocks added to the other's sqref.
* Formula1 and Formula2 read back as they were given, a list without
  quotes, a formula with its = and its relative references moved to the
  cell asked about; a date reads back as m/d/yyyy and a time as
  h:mm:ss AM/PM. The file keeps a list in quotes, a formula without its =
  relative to the rule's first cell, and a date or a time as its serial.
* Without a rule every property is error 1004 but InputTitle and
  ErrorTitle, which read "", and Value, which is True.
* Value says whether the cell's value passes: a blank cell passes while
  IgnoreBlank is on; a list of items compares the value, spaces round it
  dropped, case and all, and a list from cells compares with case
  ignored, 2 matching "2"; a whole number, a decimal, a date or a time is
  a number within the bounds, TRUE counting as 1; a text length is the
  length of the text the value shows as; a custom formula has to come to
  TRUE where the cell is.
* Delete takes the cells out of their rules, a rule splitting round them;
  Clear does too, ClearContents does not. A copy takes the rule of the
  cells copied along.
* A list read from another sheet is kept in the worksheet's x14
  extension; the model reads such a rule and does not make one.
* Inserting or deleting cells, rows or columns carries each rule's blocks
  as conditional formats carry theirs, cut into strips round the moved
  cells, and its formulas as references move, on every sheet: a list on
  Sheet1 reading Sheet2 follows Sheet2's edits. A rule the edit reshapes,
  or whose formula's reference it resizes, goes last among the sheet's
  rules, and rules it leaves alike join into the first. Rules in the x14
  extension are written back there as Excel writes them
  (scripts/measure_x14_validation_shifts.py; planned has the details).
  Reporting itself: a shift whose blocks would come together, a delete
  moving as many cells as it deletes in one line of a block across, or in
  one wider than two columns, a block reading cells relative to itself whose first cell
  moves elsewhere, and an edit that would empty the x14 extension.
"""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import Area, parse_area
from pyopenvba._xml import attributes, escape, escape_text, unescape
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBADate, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from collections.abc import Callable

    from pyopenvba.apps.excel._conditional_edits import Edit
    from pyopenvba.apps.excel._model import Worksheet

#: xlDVType, by its number, as a file names it.
TYPES: Final = ("none", "whole", "decimal", "list", "date", "time", "textLength", "custom")
WHOLE, DECIMAL, LIST, DATE, TIME, TEXT_LENGTH, CUSTOM = 1, 2, 3, 4, 5, 6, 7
#: xlDVAlertStyle, from 1.
ALERTS: Final = ("stop", "warning", "information")
#: XlFormatConditionOperator, from 1.
OPERATORS: Final = ("between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual",
                    "lessThanOrEqual")

_RULES = re.compile(r"<dataValidations\b[^>]*?(?:/>|>(.*?)</dataValidations>)", re.DOTALL)
_RULE = re.compile(r"<dataValidation\b[^>]*?(?:/>|>.*?</dataValidation>)", re.DOTALL)
_EXTENDED = re.compile(r"<x14:dataValidation\b[^>]*?(?:/>|>.*?</x14:dataValidation>)", re.DOTALL)
_FORMULA = re.compile(r"<(?:x14:)?formula([12])>(?:<xm:f>)?(.*?)(?:</xm:f>)?</(?:x14:)?formula\1>", re.DOTALL)
_SQREF = re.compile(r"<xm:sqref>(.*?)</xm:sqref>", re.DOTALL)
#: The worksheet children dataValidations comes before, in the schema's order.
_AFTER: Final = re.compile(
    r"<(?:hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|"
    r"cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|"
    r"webPublishItems|tableParts|extLst)\b|</worksheet>")


@dataclass
class Rule:
    """One validation rule and the blocks it covers."""

    areas: list[Area]
    kind: int = 0
    alert: int = 1
    operator: int = 1
    #: As a file keeps them, without =: '"a,b"', '$D$1:$D$3', '1', '43832', 'A1>B1'.
    formula1: str = ""
    formula2: str = ""
    ignore_blank: bool = True
    dropdown: bool = True
    show_input: bool = True
    show_error: bool = True
    input_title: str = ""
    input_message: str = ""
    error_title: str = ""
    error_message: str = ""
    uid: str = field(default="", compare=False)
    #: Kept in the worksheet's x14 extension, which the model reads and leaves as it is.
    extension: bool = field(default=False, compare=False)

    @property
    def anchor(self) -> tuple[int, int]:
        return self.areas[0].top, self.areas[0].left

    def settings(self) -> Rule:
        return replace(self, areas=[])


# --- the file ------------------------------------------------------------------------------------


def read_rules(xml: str) -> list[Rule]:
    """A worksheet's rules as its part spells them, those in its x14 extension after the rest."""
    rules: list[Rule] = []
    found = _RULES.search(xml)
    if found is not None and found.group(1):
        rules += [_rule(match.group(0), extension=False) for match in _RULE.finditer(found.group(1))]
    rules += [_rule(match.group(0), extension=True) for match in _EXTENDED.finditer(xml)]
    return [rule for rule in rules if rule.areas]


def _rule(element: str, *, extension: bool) -> Rule:
    head = attributes(element[:element.index(">") + 1].replace("<x14:dataValidation", "<dataValidation"))
    formulas = {number: unescape(text) for number, text in _FORMULA.findall(element)}
    listed = _SQREF.search(element)
    sqref = head.get("sqref") or (listed.group(1) if listed is not None else "")
    return Rule(
        areas=[parse_area(part, sheet="") for part in sqref.split()],
        kind=TYPES.index(head["type"]) if head.get("type") in TYPES else 0,
        alert=ALERTS.index(head["errorStyle"]) + 1 if head.get("errorStyle") in ALERTS else 1,
        operator=OPERATORS.index(head["operator"]) + 1 if head.get("operator") in OPERATORS else 1,
        formula1=formulas.get("1", ""), formula2=formulas.get("2", ""),
        ignore_blank=head.get("allowBlank") == "1", dropdown=head.get("showDropDown") != "1",
        show_input=head.get("showInputMessage") == "1", show_error=head.get("showErrorMessage") == "1",
        input_title=head.get("promptTitle", ""), input_message=head.get("prompt", ""),
        error_title=head.get("errorTitle", ""), error_message=head.get("error", ""), uid=head.get("xr:uid", ""),
        extension=extension)


def rules_xml(sheet: Worksheet) -> str:
    """The dataValidations element for the sheet's rules outside the x14 extension, "" for none."""
    rules = [rule for rule in sheet.validations if not rule.extension]
    if not rules:
        return ""
    return f'<dataValidations count="{len(rules)}">{"".join(_rule_xml(rule) for rule in rules)}</dataValidations>'


def _rule_xml(rule: Rule) -> str:
    """One rule as Excel writes it: a setting at its default left out, and the attributes in Excel's order."""
    parts: list[str] = []
    if rule.kind:
        parts.append(f'type="{TYPES[rule.kind]}"')
    if rule.alert != 1:
        parts.append(f'errorStyle="{ALERTS[rule.alert - 1]}"')
    if rule.operator != 1:
        parts.append(f'operator="{OPERATORS[rule.operator - 1]}"')
    if rule.ignore_blank:
        parts.append('allowBlank="1"')
    if not rule.dropdown:
        parts.append('showDropDown="1"')
    if rule.show_input:
        parts.append('showInputMessage="1"')
    if rule.show_error:
        parts.append('showErrorMessage="1"')
    for name, text in (("errorTitle", rule.error_title), ("error", rule.error_message),
                       ("promptTitle", rule.input_title), ("prompt", rule.input_message)):
        if text:
            parts.append(f'{name}="{escape(text)}"')
    parts.append(f'sqref="{" ".join(area.address(absolute=False) for area in rule.areas)}"')
    if rule.uid:
        parts.append(f'xr:uid="{rule.uid}"')
    formulas = "".join(f"<formula{number}>{escape_text(text)}</formula{number}>"
                       for number, text in ((1, rule.formula1), (2, rule.formula2)) if text)
    head = "<dataValidation " + " ".join(parts)
    return f"{head}>{formulas}</dataValidation>" if formulas else head + "/>"


def _extended_xml(rule: Rule) -> str:
    """One rule as Excel writes it in the x14 extension: the attributes of a plain rule but the range, its formulas
    in xm:f, then its range in xm:sqref (tests/fixtures/x14_validation_base.xlsx)."""
    head = _rule_xml(replace(rule, areas=[], formula1="", formula2=""))
    head = re.sub(r' sqref=""', "", head).removesuffix("/>").removesuffix(">").replace("<dataValidation", "<x14:dataValidation", 1)
    formulas = "".join(f"<x14:formula{number}><xm:f>{escape_text(text)}</xm:f></x14:formula{number}>"
                       for number, text in ((1, rule.formula1), (2, rule.formula2)) if text)
    ranges = " ".join(area.address(absolute=False) for area in rule.areas)
    return f"{head}>{formulas}<xm:sqref>{ranges}</xm:sqref></x14:dataValidation>"


def _with_extension(sheet: Worksheet, xml: str) -> str:
    """A worksheet part with the x14 extension's dataValidations element as the model has its rules."""
    rules = [rule for rule in sheet.validations if rule.extension]
    existing = re.search(r"<x14:dataValidations\b[^>]*>.*?</x14:dataValidations>", xml, re.DOTALL)
    if existing is None or not rules:
        # A planned edit refuses before it would empty the extension; a sheet without one keeps none.
        return xml
    markup = (f'<x14:dataValidations count="{len(rules)}" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
              f'{"".join(_extended_xml(rule) for rule in rules)}</x14:dataValidations>')
    return xml[:existing.start()] + markup + xml[existing.end():]


def with_rules(sheet: Worksheet, xml: str) -> str:
    """A worksheet part with its dataValidations element as the model has the rules: replaced, put in its place,
    or taken out; and the x14 extension's rules as the model has them."""
    xml = _with_extension(sheet, xml)
    markup = rules_xml(sheet)
    existing = re.search(r"<dataValidations\b[^>]*(?:/>|>.*?</dataValidations>)", xml, re.DOTALL)
    if existing is not None:
        return xml[:existing.start()] + markup + xml[existing.end():]
    following = _AFTER.search(xml)
    if not markup or following is None:
        return xml
    return xml[:following.start()] + markup + xml[following.start():]


# --- the rules of a sheet --------------------------------------------------------------------


def rule_at(sheet: Worksheet, row: int, column: int) -> Rule | None:
    return next((rule for rule in sheet.validations if any(area.contains(row, column) for area in rule.areas)), None)


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def _without(area: Area, cut: Area) -> list[Area]:
    """What is left of ``area`` once ``cut`` is taken out: the rows above it, beside it on the left and the
    right, and below it."""
    if not _meets(area, cut):
        return [area]
    left: list[Area] = []
    if area.top < cut.top:
        left.append(Area(area.top, area.left, cut.top - 1, area.right))
    middle_top, middle_bottom = max(area.top, cut.top), min(area.bottom, cut.bottom)
    if area.left < cut.left:
        left.append(Area(middle_top, area.left, middle_bottom, cut.left - 1))
    if cut.right < area.right:
        left.append(Area(middle_top, cut.right + 1, middle_bottom, area.right))
    if cut.bottom < area.bottom:
        left.append(Area(cut.bottom + 1, area.left, area.bottom, area.right))
    return left


def remove(sheet: Worksheet, areas: list[Area]) -> bool:
    """Take these cells out of every rule, a rule splitting round them; whether any rule changed."""
    changed = False
    for rule in list(sheet.validations):
        if rule.extension and any(_meets(one, cut) for one in rule.areas for cut in areas):
            raise VBAUnsupportedError("changing a validation kept in the worksheet's x14 extension is not "
                                      "implemented")
        kept = rule.areas
        for cut in areas:
            kept = [piece for one in kept for piece in _without(one, cut)]
        if kept != rule.areas:
            changed = True
            if kept:
                rule.areas = kept
            else:
                sheet.validations.remove(rule)
    if changed:
        sheet.validations_changed = True
        sheet.touched()
    return changed


def add(sheet: Worksheet, rule: Rule) -> None:
    """A new rule, joined to one the sheet has whose settings are the same where it stands."""
    for other in sheet.validations:
        if not other.extension and _same(other, rule):
            other.areas.extend(rule.areas)
            break
    else:
        if sheet.declares_revisions():
            rule.uid = "{" + str(uuid.uuid4()).upper() + "}"
        sheet.validations.append(rule)
    sheet.validations_changed = True
    sheet.touched()


def _same(one: Rule, other: Rule) -> bool:
    """Whether two rules are the same rule: their settings alike, and their formulas alike once moved to one's
    first cell."""
    moved = replace(other, formula1=_moved(other.formula1, other.anchor, one.anchor),
                    formula2=_moved(other.formula2, other.anchor, one.anchor))
    return moved.settings() == one.settings()


def _moved(formula: str, source: tuple[int, int], target: tuple[int, int]) -> str:
    """A rule's formula, relative to ``source``, written relative to ``target`` instead."""
    from pyopenvba.formula._parse import FormulaError, shift_text

    if not formula or formula.startswith('"'):
        return formula
    try:
        return shift_text("=" + formula, target[0] - source[0], target[1] - source[1])[1:]
    except FormulaError:
        return formula


def _relative(formula: str) -> bool:
    """Whether a rule's formula reads a cell relative to the rule's first cell."""
    from pyopenvba.formula._parse import FormulaError, tokenize

    if not formula or formula.startswith('"'):
        return False
    try:
        tokens = tokenize("=" + formula)
    except FormulaError:
        return True
    return any(token.kind == "ref" and re.search(r"(?<![$A-Za-z])[A-Za-z]{1,3}\$?[0-9]|(?<!\$)\b[0-9]+:", token.text)
               for token in tokens)


def _to_end(box: tuple[int, int, int, int], edit: Edit) -> bool:
    """Whether Excel puts a rule last among the sheet's rules for what the edit does to this block of it: a delete
    cutting its front or middle, a shift splitting it, or a shift inserting cells inside it; not a block that moves
    whole, grows by whole lines or by cells inserted just after it, or loses its last lines
    (tests/fixtures/x14_validation_shifts.json, validation_list_order.json)."""
    from pyopenvba.apps.excel._conditional_edits import LineEdit

    top, left, bottom, right = box
    vertical = edit.rows if isinstance(edit, LineEdit) else edit.vertical
    low, high = (top, bottom) if vertical else (left, right)
    if not isinstance(edit, LineEdit):
        across = (left, right) if vertical else (top, bottom)
        if across[1] < edit.band[0] or edit.band[1] < across[0] or high < edit.start:
            return False
        if not (edit.band[0] <= across[0] and across[1] <= edit.band[1]):
            # A delete taking all the block has in some of its lines keeps it whole, or without that edge, and keeps
            # its place (x14_validation_shifts.json A4:H4, x14_validation_shifts3.json B2:E2).
            if edit.delete and edit.start <= low and high <= edit.start + edit.count - 1:
                return False
            return True
        if not edit.delete:
            return low < edit.start <= high
    if not edit.delete:
        return False
    end = edit.start + edit.count - 1
    return edit.start <= high and low <= end < high


def _resized(before: str, after: str, edit: Edit, *, owner: str, edited: str) -> bool:
    """Whether what an edit did to a rule's formula puts the rule last.

    Through a cell shift, any reference to the edited sheet that grew, shrank, went to #REF! or that the shift cuts
    through, its text kept, does; one that only moved does not (validation_formula_order.json, x14_validation_shifts4.json
    A1:B1 right). Through whole lines, a reference to another sheet than the rule's that shrank does, one they grew
    does not (x14_validation_shifts4.json Sheet2's row 1 deleted, x14_validation_shifts5.json rows 2:5 inserted); on the
    rule's own sheet only a reference a delete starting inside it shrank does: not one the delete takes from its first
    line, one an insert grew, or one that went to #REF! (x14_validation_shifts.json, x14_validation_shifts2.json)."""
    from pyopenvba.apps.excel._conditional_edits import LineEdit
    from pyopenvba.formula._parse import FormulaError, tokenize

    if not before or before.startswith('"'):
        return False
    try:
        old, new = ([token.text for token in tokenize("=" + text) if token.kind == "ref"] for text in (before, after))
    except FormulaError:
        return False

    def on_edited(text: str) -> bool:
        place = text.rsplit("!", 1)[0].strip("'") if "!" in text else owner
        return place.lower() == edited.lower()

    lines = isinstance(edit, LineEdit)
    strict = not lines
    if not lines:
        for one in old:
            if not on_edited(one):
                continue
            try:
                box = parse_area(one.split("!")[-1], sheet="")
            except ValueError:
                continue
            along = (box.top, box.bottom) if edit.vertical else (box.left, box.right)
            across = (box.left, box.right) if edit.vertical else (box.top, box.bottom)
            if across[0] <= edit.band[1] and edit.band[0] <= across[1] and along[1] >= edit.start and not (
                    edit.band[0] <= across[0] and across[1] <= edit.band[1]):
                return True
    if before == after:
        return False
    if len(new) < len(old) and (strict or any("!" in one for one in old)):
        # A reference gone to #REF!, which reads as an error rather than a reference: last through a cell shift, and
        # through whole lines where it named another sheet (x14_validation_shifts6.json, Sheet2's rows 2:5 deleted).
        return True
    for one, other in zip(old, new):
        try:
            first = parse_area(one.split("!")[-1], sheet="")
        except ValueError:
            continue
        try:
            second = parse_area(other.split("!")[-1], sheet="")
        except ValueError:
            if strict:
                return True
            continue
        size, now = first.rows * first.columns, second.rows * second.columns
        if size == now:
            continue
        if strict:
            return True
        # Whole lines: a reference they grow keeps the rule's place; one a delete shrinks puts it last where it names
        # another sheet, or where the delete starts inside it.
        if now < size and edit.delete and ("!" in one or edit.start > (first.top if edit.rows else first.left)):
            return True
    return False


def _through(rule: Rule, edit: Edit, *, formula_changed: bool = False) -> tuple[list[Area], bool]:
    """A rule's blocks through an edit, and whether Excel puts the rule last for it; ``formula_changed`` where the edit
    changes the rule's formula too."""
    from pyopenvba.apps.excel._conditional_edits import LineEdit, range_through

    vertical = edit.rows if isinstance(edit, LineEdit) else edit.vertical
    areas: list[Area] = []
    last = False
    for area in rule.areas:
        box = (area.top, area.left, area.bottom, area.right)
        if not isinstance(edit, LineEdit) and edit.delete:
            low, high = (area.top, area.bottom) if vertical else (area.left, area.right)
            across = (area.left, area.right) if vertical else (area.top, area.bottom)
            partly = across[0] <= edit.band[1] and edit.band[0] <= across[1] and not (
                edit.band[0] <= across[0] and across[1] <= edit.band[1])
            if partly and low < edit.start <= high <= edit.start + edit.count - 1:
                # The band runs through the block and the delete takes its far edge there, so none of its cells
                # move in: Excel keeps the block whole, as a reference keeps it (x14_validation_shifts.json).
                areas.append(area)
                continue
            if partly and edit.start <= low and high <= edit.start + edit.count - 1 and \
                    across[0] < edit.band[0] and edit.band[1] < across[1]:
                # A delete taking all the block's cells in lines strictly inside it keeps the block whole, as a
                # reference keeps it, unless the edit changes the rule's formula too: then Excel makes the rule anew from
                # its cells and the block comes apart round the lines (x14_validation_shifts.json and later rounds: 13
                # kept whole; x14_validation_shifts9.json A3:D3 left: B2:B4 reading $A$1:$A$3 to B2 B4).
                if not formula_changed:
                    areas.append(area)
                    continue
                boxes, kinds, _ = range_through([box], edit)
                areas += [Area(*one) for one in boxes]
                last = True
                continue
            width = min(across[1], edit.band[1]) - max(across[0], edit.band[0]) + 1
            moved = high - (edit.start + edit.count - 1)
            if partly and width == 1 and low < edit.start and moved == edit.count:
                # A delete up in one column of a block two columns wide that moves as many cells as it deletes: Excel
                # writes the other column's last cells, as many, then the whole block down to above them
                # (validation_list_order.json D5 up: C6 C3:D5; x14_validation_shifts4.json B8 up: C9 B7:C8;
                # x14_validation_shifts6.json E3:H4 up: D5:D6 D2:E4). Moving more or fewer it cuts strips (E3 up,
                # E4:H5 up). Across, or in a wider block, it was not measured.
                if not vertical or area.right - area.left != 1:
                    raise VBAUnsupportedError("Deleting cells that move as many cells as they delete within a validation's "
                                              "block is not implemented")
                other = area.right if edit.band[0] == area.left else area.left
                areas += [Area(high - edit.count + 1, other, high, other), Area(area.top, area.left, high - edit.count, area.right)]
                last = True
                continue
        try:
            boxes, kinds, _ = range_through([box], edit, whole_blocks=True)
        except VBAUnsupportedError:
            raise VBAUnsupportedError("Inserting or deleting cells where a validation's blocks would come together is not "
                                      "implemented") from None
        # A block one cell long across the shift puts the strip after the moved cells first where the strip ahead of
        # them is a single cell and it is not (tests/fixtures/validation_pieces.json).
        thin = (area.left == area.right) if not vertical else (area.top == area.bottom)
        strips = [index for index, kind in enumerate(kinds[:-1]) if kind == "split"]
        if thin and not isinstance(edit, LineEdit) and len(strips) >= 2:
            first, second = boxes[strips[0]], boxes[strips[1]]
            if (first[0], first[1]) == (first[2], first[3]) and (second[0], second[1]) != (second[2], second[3]):
                boxes[strips[0]], boxes[strips[1]] = second, first
        areas += [Area(*one) for one in boxes]
        last = last or _to_end(box, edit)
    return areas, last


def planned(sheet: Worksheet, edit: Edit, rewrite_for: Callable[[str], Callable[[str], str]]) -> Callable[[], None]:
    """What an insert or a delete does to the validation of every sheet, worked out before anything moves; the call it
    answers makes the changes. The edited sheet's rules take their blocks through the edit as conditional formats do,
    and every sheet's rules their formulas as references move; a rule the edit cuts or splits goes last, and rules the
    edit leaves alike are joined into the first, its uid kept (tests/fixtures/x14_validation_shifts.json)."""
    from pyopenvba.apps.excel._conditional_edits import LineEdit
    from pyopenvba.formula._parse import FormulaError

    plans: list[tuple[Worksheet, list[Rule]]] = []
    for owner in sheet.book.sheets_:
        if not owner.validations:
            continue
        rewrite = rewrite_for(owner.name)
        staying: list[Rule] = []
        moving: list[Rule] = []
        changed = False
        for rule in owner.validations:
            formulas: list[str] = []
            for formula in (rule.formula1, rule.formula2):
                try:
                    formulas.append(rewrite("=" + formula)[1:] if formula and not formula.startswith('"') else formula)
                except FormulaError:
                    formulas.append(formula)
            changed_formula = formulas != [rule.formula1, rule.formula2]
            areas, last = _through(rule, edit, formula_changed=changed_formula) if owner is sheet else (rule.areas, False)
            # A formula whose reference grows, shrinks or goes, not only moves, puts the rule last too
            # (x14_validation_shifts2.json: A2 down, rows 3:6 deleted); through a cell shift not where the shift moves
            # the rule's own cells whole (x14_validation_shifts7.json: A2:D2 down), though where it reshapes them
            # (x14_validation_shifts8.json: A2:D2 left).
            whole = len(areas) == len(rule.areas) and areas != rule.areas and all(
                (one.rows, one.columns) == (other.rows, other.columns) for one, other in zip(areas, rule.areas))
            if isinstance(edit, LineEdit) or not whole:
                last = last or any(_resized(old, new, edit, owner=owner.name, edited=sheet.name) for old, new in zip(
                    (rule.formula1, rule.formula2), formulas, strict=True))
            if not areas:
                changed = True
                continue
            if areas != rule.areas and (areas[0].top, areas[0].left) != _moved_corner(rule.areas[0], rewrite) and \
                    any(_relative(one) for one in (rule.formula1, rule.formula2)):
                raise VBAUnsupportedError("Inserting or deleting cells where a validation reading cells relative to "
                                          "itself starts its blocks elsewhere is not implemented")
            made = replace(rule, areas=areas, formula1=formulas[0], formula2=formulas[1])
            changed = changed or made != rule or areas != rule.areas or last
            (moving if last else staying).append(made)
        rules: list[Rule] = []
        for rule in staying + moving:
            alike = next((one for one in rules if one.extension == rule.extension and _same(one, rule)), None)
            if alike is None:
                rules.append(rule)
            else:
                alike.areas = alike.areas + rule.areas
                changed = True
        if not changed:
            continue
        if any(rule.extension for rule in owner.validations) and not any(rule.extension for rule in rules):
            raise VBAUnsupportedError("Inserting or deleting cells that takes the last validation out of the "
                                      "worksheet's x14 extension is not implemented")
        plans.append((owner, rules))

    def commit() -> None:
        for owner, rules in plans:
            owner.validations = rules
            owner.validations_changed = True
            owner.touched()

    return commit


def _moved_corner(area: Area, rewrite: Callable[[str], str]) -> tuple[int, int] | None:
    """Where a block's first cell goes, as a reference to it goes; None where it is deleted."""
    try:
        corner = parse_area(rewrite("=" + Area(area.top, area.left, area.top, area.left).address(absolute=False))[1:],
                            sheet="")
    except ValueError:
        return None
    return corner.top, corner.left


def copied(source: Worksheet, area: Area, target: Worksheet, destination: Area) -> None:
    """A copy's paste: the destination takes the rules the copied cells have, and loses its own."""
    down, across = destination.top - area.top, destination.left - area.left
    pieces: list[tuple[Rule, Area]] = []
    for rule in source.validations:
        for one in rule.areas:
            if _meets(one, area):
                shared = Area(max(one.top, area.top), max(one.left, area.left), min(one.bottom, area.bottom),
                              min(one.right, area.right))
                pieces.append((rule, Area(shared.top + down, shared.left + across, shared.bottom + down,
                                          shared.right + across)))
    had = remove(target, [destination])
    for rule, placed in pieces:
        if rule.extension:
            raise VBAUnsupportedError("copying a validation kept in the worksheet's x14 extension is not implemented")
        made = replace(rule, areas=[placed], uid="",
                       formula1=_moved(rule.formula1, rule.anchor, (placed.top, placed.left)),
                       formula2=_moved(rule.formula2, rule.anchor, (placed.top, placed.left)))
        add(target, made)
    if had or pieces:
        target.validations_changed = True


# --- what a cell's value comes to ------------------------------------------------------------


def valid(sheet: Worksheet, rule: Rule, row: int, column: int) -> bool:
    """Whether the value in a cell passes its rule, as Validation.Value answers."""
    from pyopenvba.apps.excel._calc import from_vba
    from pyopenvba.formula._values import BLANK, ExcelError

    value = from_vba(sheet.book.calculator.value_of(sheet.name, row, column))
    if value is BLANK or value == "":
        return rule.ignore_blank
    if rule.kind == 0:
        return True
    if rule.kind == CUSTOM:
        answer = _worked_out(sheet, rule, rule.formula1, row, column)
        return answer is True or (isinstance(answer, (int, float)) and not isinstance(answer, bool) and answer != 0)
    if rule.kind == LIST:
        if rule.formula1.startswith('"'):
            items = [item.strip() for item in rule.formula1[1:-1].replace('""', '"').split(",")]
            return _text(value).strip() in items
        wanted = _text(value).lower()
        return any(_text(one).lower() == wanted for one in _listed(sheet, rule, row, column))
    if isinstance(value, ExcelError):
        return False
    if rule.kind == TEXT_LENGTH:
        measured: float = float(len(_text(value)))
    elif isinstance(value, bool):
        measured = float(value)
    elif isinstance(value, (int, float)):
        measured = float(value)
    else:
        return False
    if rule.kind == WHOLE and not measured.is_integer():
        return False
    bounds = [_worked_out(sheet, rule, formula, row, column) for formula in (rule.formula1, rule.formula2)]
    low = bounds[0] if isinstance(bounds[0], (int, float)) and not isinstance(bounds[0], bool) else None
    high = bounds[1] if isinstance(bounds[1], (int, float)) and not isinstance(bounds[1], bool) else None
    if low is None:
        return False
    operator = OPERATORS[rule.operator - 1]
    if operator in ("between", "notBetween"):
        inside = high is not None and low <= measured <= high
        return inside if operator == "between" else not inside
    return {"equal": measured == low, "notEqual": measured != low, "greaterThan": measured > low,
            "lessThan": measured < low, "greaterThanOrEqual": measured >= low,
            "lessThanOrEqual": measured <= low}[operator]


def _text(value: object) -> str:
    """The text a value shows as in a cell of General format."""
    from pyopenvba.formula._display import UndisplayableError, shown

    if isinstance(value, str):
        return value
    try:
        return shown(value, "General")[0] if isinstance(value, (bool, int, float)) else to_text(value)
    except UndisplayableError:
        return to_text(value)


def _evaluated(sheet: Worksheet, rule: Rule, formula: str, row: int, column: int) -> list[list[object]] | None:
    """A rule's formula worked out where the cell is, a block kept whole, as rows of values, a blank None; None where
    it cannot be worked out."""
    from pyopenvba.apps.excel._engine_book import model_value
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.lexer import FormulaSyntaxError
    from pyopenvba.formula._calc.values import Empty, ExcelError

    if not formula:
        return None
    try:
        return [[float(formula)]]
    except ValueError:
        pass
    calculator = sheet.book.calculator
    try:
        node = calculator.engine_book.read("=" + _moved(formula, rule.anchor, (row, column)))
    except FormulaSyntaxError:
        return None
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, row, column, array=True, today=now.date(), now=now)
    try:
        grid = context.array_of(context.formula(node))
    except ExcelError:
        return None
    return [[None if isinstance(one, Empty) else model_value(one) for one in line] for line in grid.rows]


def _worked_out(sheet: Worksheet, rule: Rule, formula: str, row: int, column: int) -> object:
    """A rule's formula as one value where the cell is: a block's first."""
    answer = _evaluated(sheet, rule, formula, row, column)
    return None if not answer or not answer[0] else answer[0][0]


def _listed(sheet: Worksheet, rule: Rule, row: int, column: int) -> list[object]:
    """The values a list rule's formula names, where the cell is."""
    answer = _evaluated(sheet, rule, rule.formula1, row, column)
    return [item for line in answer or [] for item in line if item is not None]


# --- VBA -------------------------------------------------------------------------------------


def _limited(value: object, limit: int) -> object:
    """A title's or message's text, or error 1004 past what Excel holds of it (tests/fixtures/member_audit.json)."""
    if len(to_text(value)) > limit:
        raise error(1004, "Application-defined or object-defined error")
    return value


class Validation(ExcelObject):
    """Range.Validation: the rule the range's cells have."""

    vba_type_name = "Validation"

    def __init__(self, target: Range) -> None:
        self.target = target
        self.sheet = target.sheet

    def _rule(self) -> Rule:
        """The one rule every cell of the range has; error 1004 where one has none or they differ."""
        remaining = [Area(area.top, area.left, area.bottom, area.right) for area in self.target.areas]
        found: list[Rule] = []
        for rule in self.sheet.validations:
            if any(_meets(one, area) for one in rule.areas for area in remaining):
                found.append(rule)
                for one in rule.areas:
                    remaining = [piece for area in remaining for piece in _without(area, one)]
        if remaining or not found or any(not _same(found[0], other) for other in found[1:]):
            raise error(1004, "Application-defined or object-defined error")
        return found[0]

    def _here(self) -> tuple[int, int]:
        return self.target.first.top, self.target.first.left

    def _shown(self, formula: str) -> str:
        """A rule's formula as Formula1 reads it for the range's first cell."""
        from pyopenvba.formula._display import shown

        rule = self._rule()
        if not formula:
            return ""
        if rule.kind == LIST and formula.startswith('"'):
            return formula[1:-1].replace('""', '"')
        try:
            number = float(formula)
        except ValueError:
            return "=" + _moved(formula, rule.anchor, self._here())
        if rule.kind == DATE:
            return shown(number, "m/d/yyyy", epoch_1904=self.sheet.book.epoch_1904,
                         locale=self.sheet.book.application.locale)[0]
        if rule.kind == TIME:
            return shown(number, "h:mm:ss AM/PM")[0]
        return formula

    def _stored(self, kind: int, value: object) -> str:
        """An argument to Add or Modify as the file keeps it."""
        from pyopenvba.apps.excel._model import spelled_formula
        from pyopenvba.apps.excel._typing import on_1904, typed_text
        from pyopenvba.formula._values import number_text

        if value is MISSING or value is None or value is EMPTY:
            return ""
        epoch_1904 = self.sheet.book.epoch_1904
        if isinstance(value, VBADate):
            if epoch_1904 and value.serial >= 1:
                # The bound is the workbook's serial, as a Date written to a cell is (tests/fixtures/date1904/).
                if value.serial < 1462:
                    raise VBAUnsupportedError("a validation bound before the 1904 calendar's first day is not "
                                              "implemented")
                return number_text(value.serial - 1462, formula=True)
            return number_text(value.serial, formula=True)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return number_text(float(value), formula=True)
        text = to_text(value)
        if text.startswith("="):
            top, left = self._here()
            spelled = spelled_formula(self.sheet, text, top, left)[1:]
            if "!" in spelled and kind == LIST:
                raise VBAUnsupportedError("a validation list read from another sheet, which Excel keeps in the "
                                          "worksheet's x14 extension, is not implemented")
            return spelled
        if kind == LIST:
            return '"' + text.replace('"', '""') + '"'
        found = typed_text(text)
        if epoch_1904:
            moved = on_1904(found, text)
            if moved is not found and isinstance(moved.value, str):
                raise VBAUnsupportedError("a validation bound before the 1904 calendar's first day is not "
                                          "implemented")
            found = moved
        typed = found.value
        if isinstance(typed, VBADate):
            typed = typed.serial
        if isinstance(typed, (int, float)) and not isinstance(typed, bool):
            return number_text(float(typed), formula=True)
        raise error(1004, "The validation's bound is not a number")

    def _made(self, Type: object, AlertStyle: object, Operator: object, Formula1: object,
              Formula2: object) -> Rule:
        kind = int(to_integer(Type, "Long"))
        if not 0 <= kind < len(TYPES):
            raise error(5, "Invalid procedure call or argument")
        alert = 1 if AlertStyle is MISSING else int(to_integer(AlertStyle, "Long"))
        operator = 1 if Operator is MISSING else int(to_integer(Operator, "Long"))
        rule = Rule(areas=[Area(area.top, area.left, area.bottom, area.right) for area in self.target.areas],
                    kind=kind, alert=alert, operator=operator)
        if kind:
            rule.formula1 = self._stored(kind, Formula1)
            rule.formula2 = self._stored(kind, Formula2) if kind not in (LIST, CUSTOM) else ""
        return rule

    @method
    def Add(self, Type: object = MISSING, AlertStyle: object = MISSING, Operator: object = MISSING,
            Formula1: object = MISSING, Formula2: object = MISSING) -> object:
        if Type is MISSING:
            raise error(449)
        if any(rule_at(self.sheet, row, column) is not None for row, column in self.target.positions()):
            raise error(1004, "Application-defined or object-defined error")
        add(self.sheet, self._made(Type, AlertStyle, Operator, Formula1, Formula2))
        return EMPTY

    @method
    def Modify(self, Type: object = MISSING, AlertStyle: object = MISSING, Operator: object = MISSING,
               Formula1: object = MISSING, Formula2: object = MISSING) -> object:
        rule = self._whole()
        made = self._made(rule.kind if Type is MISSING else Type, AlertStyle, Operator, Formula1, Formula2)
        rule.kind, rule.alert, rule.operator = made.kind, made.alert, made.operator
        rule.formula1, rule.formula2 = made.formula1, made.formula2
        self._changed()
        return EMPTY

    @method
    def Delete(self) -> object:
        remove(self.sheet, [Area(area.top, area.left, area.bottom, area.right) for area in self.target.areas])
        return EMPTY

    def _whole(self) -> Rule:
        """The rule the range is all of, for a change that cannot split one."""
        rule = self._rule()
        if rule.extension:
            raise VBAUnsupportedError("changing a validation kept in the worksheet's x14 extension is not "
                                      "implemented")
        covered = sorted((area.top, area.left, area.bottom, area.right) for area in rule.areas)
        asked = sorted((area.top, area.left, area.bottom, area.right) for area in self.target.areas)
        if covered != asked:
            raise VBAUnsupportedError("changing the validation of part of the cells a rule covers is not "
                                      "implemented")
        return rule

    def _changed(self) -> None:
        self.sheet.validations_changed = True
        self.sheet.touched()

    def _flag(self, name: str, value: object) -> None:
        rule = self._whole()
        setattr(rule, name, to_bool(value) if isinstance(getattr(rule, name), bool) else to_text(value))
        self._changed()

    @member
    def Type(self) -> object:
        return VBAInt(self._rule().kind, "Long")

    @member
    def AlertStyle(self) -> object:
        return VBAInt(self._rule().alert, "Long")

    @member
    def Operator(self) -> object:
        return VBAInt(self._rule().operator, "Long")

    @member
    def Formula1(self) -> object:
        return self._shown(self._rule().formula1)

    @member
    def Formula2(self) -> object:
        return self._shown(self._rule().formula2)

    @member
    def IgnoreBlank(self) -> object:
        return self._rule().ignore_blank

    @setter("IgnoreBlank")
    def _set_ignore_blank(self, value: object) -> None:
        self._flag("ignore_blank", value)

    @member
    def InCellDropdown(self) -> object:
        return self._rule().dropdown

    @setter("InCellDropdown")
    def _set_dropdown(self, value: object) -> None:
        self._flag("dropdown", value)

    @member
    def ShowInput(self) -> object:
        return self._rule().show_input

    @setter("ShowInput")
    def _set_show_input(self, value: object) -> None:
        self._flag("show_input", value)

    @member
    def IMEMode(self) -> object:
        """xlIMEModeNoControl, rule or none; a set, of any number, is passed over and no file keeps it. So Excel
        answered on a machine with no East Asian input method (tests/fixtures/object_members.json)."""
        return VBAInt(0, 'Long')

    @setter("IMEMode")
    def _set_ime_mode(self, value: object) -> None:
        to_integer(value, 'Long')

    @member
    def ShowError(self) -> object:
        return self._rule().show_error

    @setter("ShowError")
    def _set_show_error(self, value: object) -> None:
        self._flag("show_error", value)

    def _title(self, name: str) -> object:
        """InputTitle and ErrorTitle read "" where the cells have no rule, where the other properties fail."""
        try:
            rule = self._rule()
        except VBARuntimeError:
            return ""
        return getattr(rule, name)

    @member
    def InputTitle(self) -> object:
        return self._title("input_title")

    @setter("InputTitle")
    def _set_input_title(self, value: object) -> None:
        self._flag("input_title", _limited(value, 32))

    @member
    def InputMessage(self) -> object:
        return self._rule().input_message

    @setter("InputMessage")
    def _set_input_message(self, value: object) -> None:
        self._flag("input_message", _limited(value, 254))

    @member
    def ErrorTitle(self) -> object:
        return self._title("error_title")

    @setter("ErrorTitle")
    def _set_error_title(self, value: object) -> None:
        self._flag("error_title", _limited(value, 32))

    @member
    def ErrorMessage(self) -> object:
        return self._rule().error_message

    @setter("ErrorMessage")
    def _set_error_message(self, value: object) -> None:
        self._flag("error_message", _limited(value, 225))

    @member
    def Value(self) -> object:
        """Whether every cell of the range passes its rule; a cell with none passes."""
        for row, column in self.target.positions():
            rule = rule_at(self.sheet, row, column)
            if rule is not None and not valid(self.sheet, rule, row, column):
                return False
        return True

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.sheet.book.application
