"""Bounded, staged worksheet renames for scalar formulas and defined names."""
from __future__ import annotations

import re
import zlib
from itertools import pairwise
from typing import TYPE_CHECKING

from pyopenvba._a1 import parse_reference, quote_sheet
from pyopenvba._opc_xml import RELS_NS, opc_part
from pyopenvba.apps.excel._advanced_filter import attribute_values
from pyopenvba.apps.excel._charts_file import XmlParseError, xml_parser
from pyopenvba.apps.excel._names import entry_key, same_sheet_name
from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, parse, split_sheet, tokenize
from pyopenvba.interpreter._lex import tokenize as vba_tokens

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, NameEntry, Workbook, Worksheet
    from pyopenvba.powerquery._opc import Entry

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MAX_CELLS = 1_000_000
MAX_NAMES = 10_000
MAX_FORMULAS = 100_000
MAX_TEXT = 8_000_000
MAX_PARTS = 10_000
MAX_SHEETS = 256
MAX_EXPRESSION_TEXT = 32_768
MAX_TOKENS = 256
MAX_NESTING = 32
_OBSERVED = ('Change', 'Calculate', 'Activate', 'Deactivate', 'SelectionChange', 'PivotTableUpdate')
_RELATION_TYPES = {
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + suffix
    for suffix in ('officeDocument', 'worksheet', 'styles', 'theme', 'sharedStrings', 'calcChain', 'extended-properties')
} | {
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
}
_RESERVED = {'_xlnm.Print_Area', '_xlnm.Print_Titles', '_xlnm.Criteria', '_xlnm.Extract', '_xlnm._FilterDatabase'}


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Renaming a worksheet ' + detail + ' is not implemented')


class Budget:
    def __init__(self) -> None:
        self.remaining = MAX_TEXT

    def text(self, value: str) -> None:
        self.remaining -= len(value)
        if self.remaining < 0:
            raise unsupported('beyond the supported text budget')


def _handlers(book: Workbook, budget: Budget) -> None:
    """Do not infer event behavior from the events-disabled native controls."""
    app = book.application
    if not app.enable_events:
        return
    interpreter = app.interpreter
    if interpreter is not None:
        if len(interpreter.event_sinks) > 256:
            raise unsupported('beyond the supported event-observer count')
        for sheet in book.sheets_:
            if any(interpreter.listens_for_event(sheet, event) for event in _OBSERVED):
                raise unsupported('with enabled worksheet event observers')
        if any(interpreter.listens_for_event(owner, 'Sheet' + event)
               for owner in (book, app) for event in _OBSERVED):
            raise unsupported('with enabled workbook or application event observers')
    document_names = {'worksheet_' + event.casefold() for event in _OBSERVED}
    document_names.update('workbook_sheet' + event.casefold() for event in _OBSERVED)
    for owner in (book, *book.sheets_):
        document = owner.vba_document
        if document is not None and any(document.module.procedure(name) is not None for name in document_names):
            raise unsupported('with enabled document event observers')
    if len(book.pending_document_sources) > 1000:
        raise unsupported('beyond the supported pending-document count')
    sources = list(book.pending_document_sources.values())
    from pyopenvba.apps.excel._vbide import VBProject

    if isinstance(book.vbide_project, VBProject):
        if len(book.vbide_project.components.entries) > 1000:
            raise unsupported('beyond the supported project-component count')
        sources.extend(entry.source for entry in book.vbide_project.components.entries
                       if entry.kind == 'document' and entry.pending)
    if len(sources) > 1000:
        raise unsupported('beyond the supported pending-document count')
    for source in sources:
        if len(source) > MAX_EXPRESSION_TEXT:
            raise unsupported('beyond the supported pending-document text length')
        budget.text(source)
        try:
            tokens = vba_tokens(source)
        except VBACompileError:
            raise unsupported('with unrecognized pending document source') from None
        # The interpreter's lexer removes logical continuations and comments;
        # header names cannot hide across physical lines or inside strings.
        if any(token.kind == 'ident' and token.lower in ('sub', 'function')
               and following.kind in ('ident', 'bracket') and following.lower in document_names
               for token, following in pairwise(tokens)):
            raise unsupported('with enabled pending document event observers')


def metadata_bytes(entry: Entry, budget: Budget) -> bytes:
    """Bound actual inflated metadata, including parts not loaded by the model."""
    if entry.uncompressed_size > budget.remaining or len(entry.body) > budget.remaining:
        raise unsupported('beyond the supported package text budget')
    try:
        if entry.method == 0:
            raw = entry.body
        elif entry.method == 8:
            inflater = zlib.decompressobj(-15)
            raw = inflater.decompress(entry.body, budget.remaining + 1)
            if len(raw) > budget.remaining or not inflater.eof or inflater.unused_data:
                raise unsupported('beyond the supported or complete package metadata')
        else:
            raise unsupported('with unrecognized metadata compression')
    except zlib.error:
        raise unsupported('with malformed compressed metadata') from None
    budget.remaining -= len(raw)
    if budget.remaining < 0:
        raise unsupported('beyond the supported package text budget')
    return raw


def _workbook_names(raw: bytes) -> None:
    """Refuse name metadata the existing lexical name reader cannot own."""
    parser = xml_parser()
    path: list[str] = []
    scopes: list[int] = []
    sheets = 0
    nodes = 0
    names = 0

    def namespace(_prefix: str | None, uri: str | None) -> None:
        if uri is not None and len(uri) > 4096:
            raise unsupported('with oversized namespace metadata')

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal sheets, nodes, names
        nodes += 1
        if nodes > 100_000 or len(path) >= 64:
            raise unsupported('beyond the supported metadata node or depth count')
        path.append(name)
        if path == [NS + '}workbook', NS + '}sheets', NS + '}sheet']:
            sheets += 1
        local = name.rsplit('}', 1)[-1]
        if local in ('definedName', 'definedNames'):
            expected = [NS + '}workbook', NS + '}definedNames']
            if local == 'definedName':
                expected.append(NS + '}definedName')
            if path != expected:
                raise unsupported('with nested or foreign name metadata')
        if local == 'definedName':
            if raw[parser.CurrentByteIndex:parser.CurrentByteIndex + 12] != b'<definedName':
                raise unsupported('with qualified name element spelling')
            names += 1
            bare = attrs.get('name', '')
            if names > MAX_NAMES or not bare or len(bare) > 255 or '!' in bare:
                raise unsupported('with unrecognized stored name labels')
            if 'localSheetId' in attrs:
                scope = attrs['localSheetId']
                if re.fullmatch(r'[0-9]{1,10}', scope) is None:
                    raise unsupported('with malformed stored name scopes')
                scopes.append(int(scope))

    def end(_name: str) -> None:
        path.pop()

    parser.StartNamespaceDeclHandler = namespace
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(raw, True)
    from pyopenvba.apps.excel._io import defined_name_spans

    workbook = opc_part(raw, NS, 'workbook')
    containers = [(start, end) for tag, _attrs, start, end in workbook.children
                  if tag == NS + '}definedNames']
    # The existing writer recognizes only these exact container spellings. A
    # comment/CDATA lookalike must not become its replacement target on save.
    lexical = [match.span() for match in re.finditer(
        rb'<definedNames>.*?</definedNames>|<definedNames/>', raw, re.DOTALL)]
    actual = opc_part(raw, NS, 'workbook', container_path=('definedNames',))
    spans = [] if actual is None else [(start, end) for _tag, _attrs, start, end in actual.children]
    if containers != lexical or spans != defined_name_spans(raw):
        raise unsupported('with unowned lexical name metadata')
    if any(scope >= sheets for scope in scopes):
        raise unsupported('with out-of-range stored name scopes')


def _package(book: Workbook, budget: Budget, old: str, wanted: str) -> None:
    package = book.package
    if package is None:
        return
    if len(package.entries) > MAX_PARTS:
        raise unsupported('beyond the supported package-part budget')
    forbidden = ('xl/charts/', 'xl/drawings/', 'xl/pivot', 'xl/tables/', 'xl/externalLinks/',
                 'xl/queryTables/', 'xl/activeX/', 'xl/embeddings/', 'xl/ctrlProps/')
    worksheet_parts = {sheet.part_name for sheet in book.sheets_ if sheet.part_name}
    for entry in package.entries:
        if entry.name.startswith(forbidden) or entry.name == 'xl/connections.xml':
            raise unsupported('with unmeasured reference-bearing package parts')
        if entry.name.endswith('.rels'):
            relations = opc_part(metadata_bytes(entry, budget), RELS_NS, 'Relationships')
            for tag, attrs, _start, _end in relations.children:
                if (tag != RELS_NS + '}Relationship' or attrs.get('Type') not in _RELATION_TYPES
                        or attrs.get('TargetMode', '').casefold() not in ('', 'internal')):
                    raise unsupported('with unmeasured or external package relationships')
        elif (entry.name == 'xl/workbook.xml' or entry.name in worksheet_parts
              or (entry.name.startswith('xl/worksheets/') and entry.name.endswith('.xml'))):
            raw = metadata_bytes(entry, budget)
            if entry.name == 'xl/workbook.xml':
                _workbook_names(raw)
                names = opc_part(raw, NS, 'workbook', container_path=('definedNames',))
                if names is not None:
                    for tag, attrs, _start, _end in names.children:
                        if tag != NS + '}definedName' or (attrs.get('name', '').startswith('_xlnm') and attrs.get('name') not in _RESERVED):
                            raise unsupported('with unmodeled reserved names')
            else:
                part = opc_part(raw, NS, 'worksheet')
                from pyopenvba.apps.excel._worksheet_rename_arrays import stored

                stored(raw)
                from pyopenvba.apps.excel._worksheet_rename_secondary import worksheet

                worksheet(raw, old, wanted, budget)
                guarded = {'hyperlinks', 'drawing', 'legacyDrawing',
                           'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'tableParts', 'extLst',
                           'pivotTableParts', 'dataConsolidate', 'scenarios', 'customSheetViews', 'cellWatches', 'smartTags'}
                if any(tag.rsplit('}', 1)[-1] in guarded or not tag.startswith(NS + '}')
                       for tag, _attrs, _start, _end in part.children):
                    raise unsupported('with unmeasured worksheet reference metadata')


def _formula(text: str, old: str, wanted: str, budget: Budget, *, name_expression: bool = False) -> str:
    """Rewrite parsed local qualifiers in linear source/output character work."""
    if len(text) > MAX_EXPRESSION_TEXT:
        raise unsupported('beyond the supported expression text length')
    try:
        tokens = tokenize(text)
        if len(tokens) > MAX_TOKENS:
            raise unsupported('beyond the supported formula-token count')
        depth = 0
        powers = 0
        for token in tokens:
            depth += int(token.kind in ('open', 'lbrace')) - int(token.kind in ('close', 'rbrace'))
            powers += int(token.kind == 'op' and token.text == '^')
            if depth > MAX_NESTING or powers > MAX_NESTING:
                raise unsupported('beyond the supported formula nesting')
            owner, _reference = split_sheet(token.text) if token.kind != 'text' else ('', '')
            if token.kind == 'structured' or (owner and (':' in owner or '[' in owner or ']' in owner)):
                raise unsupported('with structured, external or three-dimensional references')
            if token.kind == 'op' and token.text == ':':
                raise unsupported('with unmeasured reference operators')
            if owner and token.kind not in ('ref', 'name'):
                raise unsupported('with qualified errors or unrecognized qualifiers')
        references = False
        if name_expression:
            try:
                parse_reference(text.removeprefix('='))
                references = True
            except ValueError:
                pass
        if not references:
            parse(text)
    except FormulaError:
        raise unsupported('with an unrecognized scalar formula or name expression') from None
    changes: list[tuple[int, int, str]] = []
    length = len(text)
    for token in tokens:
        if token.kind not in ('ref', 'name'):
            continue
        owner, reference = split_sheet(token.text)
        if owner.casefold() == old.casefold():
            replacement = quote_sheet(wanted) + '!' + reference
            changes.append((token.at, token.at + len(token.text), replacement))
            length += len(replacement) - len(token.text)
    budget.remaining -= length
    if budget.remaining < 0:
        raise unsupported('beyond the supported staged text budget')
    parts: list[str] = []
    previous = 0
    for start, finish, replacement in changes:
        parts.extend((text[previous:start], replacement))
        previous = finish
    parts.append(text[previous:])
    return ''.join(parts)


def formula(text: str, old: str, wanted: str, budget: Budget) -> str:
    """Shared scalar/name syntax admission for secondary reference inventories."""
    return _formula(text, old, wanted, budget, name_expression=True)


def rename(sheet: Worksheet, wanted: str) -> None:
    """Admit every proposal before publishing any text, scope or graph change."""
    book = sheet.book
    if sheet not in book.sheets_:
        raise unsupported('through a removed worksheet proxy')
    if wanted == sheet.name:
        book.saved = False
        return
    if book.application.udf_callers:
        raise unsupported('from a worksheet function')
    if (len(book.sheets_) > MAX_SHEETS or sum(len(owner.cells_) for owner in book.sheets_) > MAX_CELLS
            or len(book.names_.entries) > MAX_NAMES):
        raise unsupported('beyond the supported cell or name count')
    if book.calculation_in_progress:
        raise unsupported('during an active calculation')
    budget = Budget()
    _handlers(book, budget)
    if any(owner.spills or any(cell.dynamic or cell.spilled_from is not None for cell in owner.cells_.values())
           for owner in book.sheets_):
        raise unsupported('with dynamic spill storage')
    if any(owner.tables for owner in book.sheets_):
        raise unsupported('with table references')
    from pyopenvba.apps.excel._worksheet_rename_secondary import admit

    try:
        admit(book, sheet.name, wanted, budget)
        _package(book, budget, sheet.name, wanted)
    except (ValueError, XmlParseError):
        raise unsupported('with malformed persisted XML metadata') from None
    cells: list[tuple[Worksheet, Cell, str]] = []
    total = 0
    for owner in book.sheets_:
        for cell in owner.cells_.values():
            if cell.formula:
                total += 1
                if total > MAX_FORMULAS:
                    raise unsupported('beyond the supported formula count')
                if len(cell.formula) > MAX_EXPRESSION_TEXT:
                    raise unsupported('beyond the supported expression text length')
                budget.text(cell.formula)
    for entry in book.names_.entries:
        if len(entry.name) > 512:
            raise unsupported('beyond the supported name-label length')
        budget.text(entry.name)
        if len(entry.refers_to) > MAX_EXPRESSION_TEXT:
            raise unsupported('beyond the supported expression text length')
        budget.text(entry.refers_to)
        budget.text(entry.attributes)
    from pyopenvba.apps.excel._worksheet_rename_arrays import admit as storage_admit

    shared_updates = storage_admit(book, sheet, wanted, budget)
    for owner in book.sheets_:
        for cell in owner.cells_.values():
            if cell.formula:
                proposed = _formula(cell.formula, sheet.name, wanted, budget)
                cells.append((owner, cell, proposed))
    names: list[tuple[NameEntry, str, str]] = []
    seen: set[tuple[int, bytes | str, bytes | str]] = set()
    for entry in book.names_.entries:
        owner, bare = split_sheet(entry.name)
        if owner and not any(same_sheet_name(owner, one.name) for one in book.sheets_):
            raise unsupported('with an unregistered defined-name scope')
        name = quote_sheet(wanted) + '!' + bare if same_sheet_name(owner, sheet.name) else entry.name
        if entry_key(name) in seen:
            raise unsupported('with conflicting defined names')
        seen.add(entry_key(name))
        # Validate opaque attribute syntax before the names writer can be reached.
        attribute_values(entry.attributes)
        names.append((entry, name, _formula(entry.refers_to, sheet.name, wanted, budget, name_expression=True)))
    # Rebuild is callback-free for this admitted scalar-only population. Its
    # invalidation must not discard manual cached values or invent save dirt.
    dirty = [(cell, cell.save_dirty if cell.save_dirty is not None else cell.stale) for _, cell, _ in cells]
    for owner, cell, formula in cells:
        if cell.formula != formula:
            cell.formula = formula
            owner.touched()
    for entry, name, refers in names:
        entry.name, entry.refers_to = name, refers
    for key, area in shared_updates:
        sheet.shared_groups[key] = area
    sheet.name = wanted
    book.names_.changed = bool(names) or book.names_.changed
    book.saved = False
    calculator = book.calculator
    calculator.rebuild()
    if not calculator.automatic:
        for cell, saved_dirty in dirty:
            cell.save_dirty = saved_dirty
