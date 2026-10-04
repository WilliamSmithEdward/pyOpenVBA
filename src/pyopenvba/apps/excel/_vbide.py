"""In-memory VBA project components and editable source buffers.

Editing is independent of compilation: the VBE accepts incomplete source.
Pending source is compiled when execution enters the owning project.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._runtime import Slot
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_integer, to_text
from pyopenvba.vba import VBAModule, VBAModuleKind, VBAProject, VBE_REFUSED_MODULE_NAMES, split_attribute_header

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook

INVALID_ARGUMENT = -2147024809
PROCEDURE = re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)', re.I)
END_PROCEDURE = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)
REFUSED_NAMES = VBE_REFUSED_MODULE_NAMES | frozenset('''
if then else end for next do loop while wend sub function get let set dim const
public private friend static option select case with as new nothing true false
null empty boolean byte integer long string variant date debug me print rem
optional byref byval call type enum redim preserve on resume goto exit and or not
xor is like mod global
'''.split())


def body_lines(source: str) -> list[str]:
    body = split_attribute_header(source)[1].replace('\r\n', '\n').replace('\r', '\n')
    return body.split('\n') if body else []


def valid_name(name: str) -> None:
    if not name or len(name) > 31 or not name[0].isalpha() or not name.isidentifier() or name.casefold() in REFUSED_NAMES:
        raise error(50132, 'Invalid component or project name')


class VBProject(VBAObject):
    vba_type_name = 'VBProject'
    vba_library = 'vbide'

    def __init__(self, book: Workbook) -> None:
        self.book = book
        self.name = 'VBAProject'
        self.original_name = self.name
        self.components = VBComponents(self)
        self.renames: dict[str, str] = {}
        self.protected = False
        if book.package is not None and book.package.has('xl/vbaProject.bin'):
            from pyopenvba.apps.excel._vba_save import MemoryExcelFile

            with MemoryExcelFile(book.package.serialize()) as host:
                project = host.vba_project()
                self.name = project.name or self.name
                self.original_name = self.name
                self.protected = bool(project.protection and project.protection.has_password)
                kinds = host.component_kinds()
                documents = {book.code_name.casefold(), *(sheet.code_name.casefold() for sheet in book.sheets_)}
                for module in project.modules:
                    kind = kinds.get(module.name.casefold(), 'document' if module.name.casefold() in documents else 'standard' if module.kind.value == 0x21 else 'class')
                    self.components.entries.append(VBComponent(self, module.name, kind, module.source))
        self.components.synchronize()

    @member
    def Name(self) -> object:
        return self.name

    @setter('Name')
    def set_name(self, value: object) -> None:
        name = to_text(value)
        valid_name(name)
        if name.casefold() != self.name.casefold():
            self.name = name
            self.book.saved = False

    @member
    def VBComponents(self) -> object:
        if self.protected:
            raise error(50289, 'Project is protected')
        self.components.synchronize()
        return self.components

    @member
    def Type(self) -> object:
        return VBAInt(100, 'Long')

    @member
    def Mode(self) -> object:
        return VBAInt(0 if self.book.application.executing_projects else 2, 'Long')

    @member
    def FileName(self) -> object:
        return str(self.book.FullName()) if self.book.path else ''

    @member
    def Protection(self) -> object:
        return VBAInt(int(self.protected), 'Long')

    def compile_pending(self) -> None:
        from pyopenvba.apps.excel._projects import attach_project
        runtime = attach_project(self.book)
        if runtime is None:
            return
        pending = [entry for entry in self.components.entries if entry.pending]
        if pending:
            # Install the whole parsed set before binding documents, so
            # document declarations can reference classes added afterward.
            project = VBAProject(modules=[VBAModule(entry.name, entry.name, entry.source,
                                                    VBAModuleKind.standard if entry.kind == 'standard' else VBAModuleKind.other)
                                          for entry in pending])
            runtime.load_project(project)
        for entry in pending:
            entry.pending = False


class VBComponents(VBACollection):
    vba_type_name = 'VBComponents'
    vba_library = 'vbide'

    def __init__(self, project: VBProject) -> None:
        self.project = project
        self.entries: list[VBComponent] = []

    def synchronize(self) -> None:
        book = self.project.book
        names = {entry.name.casefold(): entry for entry in self.entries}
        runtime = book.project_runtime
        if not self.entries:
            for name in [book.code_name, *(sheet.code_name for sheet in book.sheets_)]:
                if name:
                    entry = VBComponent(self.project, name, 'document', '')
                    self.entries.append(entry)
                    names[name.casefold()] = entry
        if runtime is not None:
            for module in runtime.modules.values():
                entry = names.get(module.name.casefold())
                if entry is None:
                    entry = VBComponent(self.project, module.name, module.parsed.kind, module.parsed.source)
                    self.entries.append(entry)
                    names[entry.name.casefold()] = entry
                elif not entry.pending:
                    entry.source = module.parsed.source
        for name in [book.code_name, *(sheet.code_name for sheet in book.sheets_)]:
            if name and name.casefold() not in names:
                entry = VBComponent(self.project, name, 'document', '')
                self.entries.append(entry)
                names[name.casefold()] = entry
        for name, source in book.pending_document_sources.items():
            entry = names.get(name.casefold())
            if entry is not None and not entry.pending:
                entry.source = source
        deleted = {name.casefold() for name in book.deleted_document_names}
        self.entries[:] = [entry for entry in self.entries if entry.name.casefold() not in deleted]

    def vba_items(self) -> list[object]:
        self.synchronize()
        return list(self.entries)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for entry in self.entries:
                if entry.name.casefold() == index.casefold():
                    return entry
            raise error(9)
        return super().vba_lookup(index, items)

    @member
    def Parent(self) -> object:
        return self.project

    @method
    def Add(self, ComponentType: object) -> object:
        number = to_integer(ComponentType, 'Long')
        if number not in (1, 2):
            if number == 3:
                raise VBAUnsupportedError('VBComponents.Add(UserForm) is not implemented')
            raise error(440)
        self.synchronize()
        prefix = 'Module' if number == 1 else 'Class'
        names = {entry.name.casefold() for entry in self.entries}
        counter = 1
        while f'{prefix}{counter}'.casefold() in names:
            counter += 1
        entry = VBComponent(self.project, f'{prefix}{counter}', 'standard' if number == 1 else 'class', '')
        entry.pending = True
        self.entries.append(entry)
        self.project.book.deleted_document_names = {name for name in self.project.book.deleted_document_names if name.casefold() != entry.name.casefold()}
        self.project.book.saved = False
        return entry

    @method
    def Remove(self, VBComponent: object) -> object:
        component = require_component(VBComponent)
        if component not in self.entries or component.kind == 'document':
            raise error(5)
        self.entries.remove(component)
        component.invalidated = True
        book = self.project.book
        if book.project_runtime is not None:
            book.project_runtime.modules.pop(component.name.casefold(), None)
        book.deleted_document_names.add(component.name)
        book.saved = False
        return EMPTY


class VBComponent(VBAObject):
    vba_type_name = 'VBComponent'
    vba_library = 'vbide'

    def __init__(self, project: VBProject, name: str, kind: str, source: str) -> None:
        self.project = project
        self.name = name
        self.kind = kind
        self.source = source
        self.pending = False
        self.invalidated = False
        self.code_module = CodeModule(self)

    @member
    def Name(self) -> object:
        return self.name

    @setter('Name')
    def set_name(self, value: object) -> None:
        name = to_text(value)
        valid_name(name)
        if name.casefold() == self.name.casefold():
            return
        if name.casefold() == 'excel' or any(entry.name.casefold() == name.casefold() for entry in self.project.components.entries):
            raise error(32813)
        book = self.project.book
        old = self.name
        self.project.renames[old] = name
        if book.project_runtime is not None:
            runtime = book.project_runtime.modules.pop(old.casefold(), None)
            if runtime is not None:
                runtime.name = runtime.parsed.name = name
                runtime.parsed.source = re.sub(r'(?m)^(Attribute VB_Name\s*=\s*)"[^"]*"', lambda match: match.group(1) + '"' + name + '"', runtime.parsed.source)
                book.project_runtime.modules[name.casefold()] = runtime
        if self.kind == 'document':
            if book.code_name.casefold() == old.casefold():
                book.code_name = name
            for sheet in book.sheets_:
                if sheet.code_name.casefold() == old.casefold():
                    sheet.code_name = name
            book.used_code_names.add(old.casefold())
            book.used_code_names.add(name.casefold())
        if old in book.pending_document_sources:
            book.pending_document_sources[name] = book.pending_document_sources.pop(old)
        self.name = name
        self.source = re.sub(r'(?m)^(Attribute VB_Name\s*=\s*)"[^"]*"', lambda match: match.group(1) + '"' + name + '"', self.source)
        book.saved = False

    @member
    def Type(self) -> object:
        return VBAInt({'standard': 1, 'class': 2, 'form': 3, 'document': 100}[self.kind], 'Long')

    @member
    def CodeModule(self) -> object:
        return self.code_module

    @member
    def Collection(self) -> object:
        return self.project.components

class CodeModule(VBAObject):
    vba_type_name = 'CodeModule'
    vba_library = 'vbide'

    def __init__(self, component: VBComponent) -> None:
        self.component = component

    def lines(self) -> list[str]:
        return body_lines(self.component.source)

    def edit(self, lines: list[str]) -> None:
        header = split_attribute_header(self.component.source)[0]
        source = header + '\r\n'.join(lines)
        if source != self.component.source:
            self.component.source = source
            self.component.pending = True
            self.component.project.book.saved = False

    @member
    def Parent(self) -> object:
        return self.component

    @member
    def CountOfLines(self) -> object:
        return VBAInt(len(self.lines()), 'Long')

    @member
    def Lines(self, StartLine: object, Count: object) -> object:
        start, count = int(to_integer(StartLine, 'Long')), int(to_integer(Count, 'Long'))
        if start <= 0 or count <= 0:
            raise error(INVALID_ARGUMENT)
        lines = self.lines()
        result = '\r\n'.join(lines[start - 1:start - 1 + count])
        return result + ('\r\n' if result and start - 1 + count > len(lines) else '')

    @method
    def InsertLines(self, Line: object, String: object) -> object:
        line = int(to_integer(Line, 'Long'))
        if line <= 0:
            raise error(INVALID_ARGUMENT)
        lines = self.lines()
        position = min(line - 1, len(lines))
        added = to_text(String).replace('\r\n', '\n').replace('\r', '\n').split('\n')
        self.edit(lines[:position] + added + lines[position:])
        return EMPTY

    @method
    def AddFromString(self, String: object) -> object:
        text = to_text(String)
        if text:
            self.InsertLines(int(self.CountOfDeclarationLines()) + 1, text)
        return EMPTY

    @method
    def DeleteLines(self, StartLine: object, Count: object = MISSING) -> object:
        start = int(to_integer(StartLine, 'Long'))
        count = 1 if Count is MISSING else int(to_integer(Count, 'Long'))
        if start <= 0 or count < 0:
            raise error(5)
        lines = self.lines()
        if start > len(lines) or start - 1 + count > len(lines):
            raise error(INVALID_ARGUMENT)
        self.edit(lines[:start - 1] + lines[start - 1 + count:])
        return EMPTY

    @method
    def ReplaceLine(self, Line: object, String: object) -> object:
        line = int(to_integer(Line, 'Long'))
        lines = self.lines()
        if line <= 0 or line > len(lines):
            raise error(INVALID_ARGUMENT)
        self.edit(lines[:line - 1] + to_text(String).replace('\r\n', '\n').replace('\r', '\n').split('\n') + lines[line:])
        return EMPTY

    def procedures(self) -> list[tuple[str, int, int, int, int]]:
        lines = self.lines()
        result: list[tuple[str, int, int, int, int]] = []
        previous_end = 0
        position = 0
        while position < len(lines):
            found = PROCEDURE.match(lines[position])
            if found is None:
                position += 1
                continue
            body = position + 1
            start = body
            while start > previous_end + 1 and (not lines[start - 2].strip() or lines[start - 2].lstrip().startswith("'")):
                start -= 1
            kind = {'get': 3, 'let': 1, 'set': 2}.get((found.group(2) or '').lower(), 0)
            end = position + 1
            while end < len(lines) and not END_PROCEDURE.match(lines[end]):
                if PROCEDURE.match(lines[end]):
                    break
                end += 1
            end = min(end + 1, len(lines))
            result.append((found.group(3), kind, start, body, end))
            previous_end = end
            position = end
        if result and all(not line.strip() or line.lstrip().startswith("'") for line in lines[result[-1][4]:]):
            last = result[-1]
            result[-1] = (*last[:4], len(lines))
        return result

    @member
    def CountOfDeclarationLines(self) -> object:
        procedures = self.procedures()
        return VBAInt(procedures[0][2] - 1 if procedures else len(self.lines()), 'Long')

    def procedure(self, name: object, kind: object) -> tuple[str, int, int, int, int]:
        for entry in self.procedures():
            if entry[0].casefold() == to_text(name).casefold() and entry[1] == int(to_integer(kind, 'Long')):
                return entry
        raise error(35)

    @member
    def ProcStartLine(self, ProcName: object, ProcKind: object) -> object:
        return VBAInt(self.procedure(ProcName, ProcKind)[2], 'Long')

    @member
    def ProcBodyLine(self, ProcName: object, ProcKind: object) -> object:
        return VBAInt(self.procedure(ProcName, ProcKind)[3], 'Long')

    @member
    def ProcCountLines(self, ProcName: object, ProcKind: object) -> object:
        entry = self.procedure(ProcName, ProcKind)
        return VBAInt(entry[4] - entry[2] + 1, 'Long')

    @member
    def ProcOfLine(self, Line: object, ProcKind: object) -> object:
        line = int(to_integer(Line, 'Long'))
        for entry in self.procedures():
            if entry[2] <= line <= entry[4]:
                if isinstance(ProcKind, Slot):
                    ProcKind.set(VBAInt(entry[1], 'Long'))
                return entry[0]
        return ''


def project_for(book: Workbook) -> VBProject:
    from pyopenvba.apps.excel._documents import materialize

    materialize(book)
    if not isinstance(book.vbide_project, VBProject):
        book.vbide_project = VBProject(book)
    return book.vbide_project


def require_component(value: object) -> VBComponent:
    if not isinstance(value, VBComponent):
        raise error(5)
    return value
