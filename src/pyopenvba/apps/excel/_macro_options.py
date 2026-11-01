"""Application.MacroOptions: a macro's description, help id, shortcut key and function category, which Excel keeps as
attribute lines inside the procedure.

Measured in live Excel (scripts/measure_macro_options.py, tests/fixtures/macro_options.json), reading each module back
from the workbook saved after it:

- Every call writes ``Attribute F.VB_ProcData.VB_Invoke_Func = "k\\n14"``: the shortcut key, a space for none, and
  the category, 14 (User Defined) until given. A later call keeps both unless it gives them again; HasShortcutKey
  False takes no key away.
- A description goes in VB_Description, its quotes doubled, ahead of the rest; an empty one takes it out.
  HelpContextID goes in VB_HelpID, between the two.
- The lines follow the procedure's first line, which the model takes whole where it goes on with `` _``. The editor
  keeps them out of sight: CountOfLines leaves them out, and lines put in or taken out around them leave them with
  their procedure (see _vbide.hidden_attributes).
- A macro the active workbook's modules do not hold, a shortcut key that is not one letter and a category of 40 are
  error 1004.

A category given by name, a category past 14, ArgumentDescriptions, StatusBar, HelpFile, HasMenu, MenuText and a
macro named with its workbook each report themselves unsupported: where Excel keeps them was not found, or they were
not measured.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application
    from pyopenvba.apps.excel._projects import ExcelInterpreter
    from pyopenvba.apps.excel._vbide import VBComponent

FAILED = "Method 'MacroOptions' of object '_Application' failed"
#: The attributes MacroOptions writes, in the order Excel writes them.
ORDER = ('VB_Description', 'VB_HelpID', 'VB_ProcData.VB_Invoke_Func')


def _header(name: str) -> re.Pattern[str]:
    return re.compile(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+'
                      + re.escape(name) + r'\b', re.IGNORECASE)


def macro_options(application: Application, macro: object, description: object, has_shortcut_key: object,
                  shortcut_key: object, category: object, help_context_id: object,
                  unmeasured: dict[str, object]) -> object:
    from pyopenvba.apps.excel._projects import attach_project

    for name, given in unmeasured.items():
        if given is not MISSING:
            raise VBAUnsupportedError(f"MacroOptions' {name} is not implemented")
    text = to_text(macro)
    if '!' in text:
        raise VBAUnsupportedError("MacroOptions on a macro named with its workbook is not implemented")
    module_name, _, procedure = text.rpartition('.')
    book = application.active_book
    project = attach_project(book) if book is not None else None
    if book is None or project is None:
        raise error(1004, FAILED)
    from pyopenvba.apps.excel._vbide import VBProject

    # The VBA project's components where the book has them, its modules waiting to compile among them; else the
    # modules the interpreter runs.
    modules: list[tuple[str, str, Callable[[str], None]]] = []
    if isinstance(book.vbide_project, VBProject):
        for component in book.vbide_project.components.entries:
            if component.kind == 'standard':
                modules.append((component.name, component.persisted_source(), _component_setter(component)))
    else:
        for runtime in project.modules.values():
            if runtime.parsed.kind == 'standard':
                modules.append((runtime.name, runtime.parsed.source, _runtime_setter(project, runtime.name)))
    found = None
    for name, source, put in modules:
        if module_name and name.casefold() != module_name.casefold():
            continue
        lines = source.splitlines(keepends=True)
        start = next((index for index, line in enumerate(lines) if _header(procedure).match(line)), None)
        if start is not None:
            found = put, lines, start
            break
    if found is None:
        raise error(1004, FAILED)
    put, lines, start = found
    end = start
    while lines[end].rstrip('\r\n').endswith(' _'):
        end += 1
    newline = '\r\n' if lines[start].endswith('\r\n') else '\n'
    attributes: dict[str, str] = {}
    after = end + 1
    prefix = re.compile(r'^Attribute\s+' + re.escape(procedure) + r'\.(\S+)\s*=\s*(.*)$', re.IGNORECASE)
    while after < len(lines) and (matched := prefix.match(lines[after].rstrip('\r\n'))):
        if matched[1] not in ORDER:
            raise VBAUnsupportedError(f'MacroOptions on a procedure with its own {matched[1]} is not implemented')
        attributes[matched[1]] = matched[2]
        after += 1
    key, number = ' ', 14
    if 'VB_ProcData.VB_Invoke_Func' in attributes:
        found_key = re.fullmatch(r'"(.)\\n(\d+)"', attributes['VB_ProcData.VB_Invoke_Func'])
        if found_key is None:
            raise VBAUnsupportedError('MacroOptions on a procedure with a shortcut of this form is not implemented')
        key, number = found_key[1], int(found_key[2])
    if shortcut_key is not MISSING:
        wanted = to_text(shortcut_key)
        if len(wanted) != 1 or not wanted.isascii() or not wanted.isalpha():
            raise error(1004, 'Shortcut key must be a letter.')
        key = wanted
    if has_shortcut_key is not MISSING:
        to_bool(has_shortcut_key)
    if category is not MISSING:
        if isinstance(category, str):
            raise VBAUnsupportedError('MacroOptions with a category given by name is not implemented')
        number = int(to_integer(category, 'Long'))
        if number == 40:
            raise error(1004, FAILED)
        if not 0 <= number <= 14:
            raise VBAUnsupportedError('MacroOptions with a category past 14 is not implemented')
    if description is not MISSING:
        wanted = to_text(description)
        if wanted:
            attributes['VB_Description'] = '"' + wanted.replace('"', '""') + '"'
        else:
            attributes.pop('VB_Description', None)
    if help_context_id is not MISSING:
        attributes['VB_HelpID'] = str(int(to_integer(help_context_id, 'Long')))
    attributes['VB_ProcData.VB_Invoke_Func'] = f'"{key}\\n{number}"'
    written = [f'Attribute {procedure}.{name} = {attributes[name]}{newline}' for name in ORDER if name in attributes]
    put(''.join(lines[:end + 1] + written + lines[after:]))
    book.saved = False
    return EMPTY


def _component_setter(component: VBComponent) -> Callable[[str], None]:
    def put(source: str) -> None:
        component.code_module.editor_lines = None
        component.source = source
        component.pending = True
        component.compile_invalidated = True

    return put


def _runtime_setter(project: ExcelInterpreter, name: str) -> Callable[[str], None]:
    def put(source: str) -> None:
        project.add_module(source, name=name, kind='standard')

    return put
