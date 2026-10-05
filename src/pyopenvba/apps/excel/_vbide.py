"""In-memory VBA project components and editable source buffers.

Editing is independent of compilation: the VBE accepts incomplete source.
Pending source is compiled when execution enters the owning project.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._runtime import Slot
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_integer, to_text
from pyopenvba.vba import VBAModule, VBAModuleKind, VBAProject, VBE_REFUSED_MODULE_NAMES, encode_mbcs, encoding_for_codepage, normalize_class_source, split_attribute_header

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


def editor_source_parts(source: str) -> tuple[str, str]:
    prefix = ''
    body = source
    if body.startswith('VERSION ') and ' CLASS' in body[:64]:
        found = re.search(r'(?:\r\n|\n|\r)END(?:\r\n|\n|\r)', body)
        if found is not None:
            prefix = body[:found.end()]
            body = body[found.end():]
    found = re.match(r'(?:Attribute [^\r\n]*(?:\r\n|\n|\r|$))*', body)
    assert found is not None
    return prefix + found.group(), body[found.end():]


def body_lines(source: str) -> list[str]:
    return editor_source_parts(source)[1].splitlines()


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
        self.code_page = 1252
        self.components = VBComponents(self)
        self.renames: dict[str, str] = {}
        self.protected = False
        if book.package is not None and book.package.has('xl/vbaProject.bin'):
            from pyopenvba.apps.excel._vba_save import MemoryExcelFile

            with MemoryExcelFile(book.package.serialize()) as host:
                project = host.vba_project()
                self.name = project.name or self.name
                self.original_name = self.name
                self.code_page = project.code_page
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
        if not self.book.path:
            raise error(76)
        return str(self.book.FullName())

    @member
    def Protection(self) -> object:
        return VBAInt(int(self.protected), 'Long')

    def compile_pending(self) -> None:
        from pyopenvba.apps.excel._projects import attach_project
        runtime = attach_project(self.book)
        if runtime is None:
            return
        # ReplaceLine changes saved source, but native Excel keeps a module's
        # already compiled code until an insertion/deletion/rename invalidates it.
        pending = [entry for entry in self.components.entries if entry.pending and
                   (entry.compile_invalidated or runtime.modules.get(entry.name.casefold()) not in runtime.executed_modules)]
        if pending:
            # Install the whole parsed set before binding documents, so
            # document declarations can reference classes added afterward.
            project = VBAProject(modules=[VBAModule(entry.name, entry.name, entry.source,
                                                    VBAModuleKind.standard if entry.kind == 'standard' else VBAModuleKind.other)
                                          for entry in pending])
            runtime.load_project(project)
        for entry in pending:
            entry.pending = False
            entry.compile_invalidated = False


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
                    if entry.source != module.parsed.source:
                        entry.code_module.editor_lines = None
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

    @method
    def Import(self, FileName: object) -> object:
        path = Path(to_text(FileName))
        text = read_source_file(path, self.project.code_page)
        if path.suffix.casefold() == '.frm':
            raise VBAUnsupportedError('VBComponents.Import(UserForm) is not implemented')
        kind = 'class' if path.suffix.casefold() == '.cls' else 'standard'
        if kind == 'class':
            text = normalize_class_source(text)
        found = re.search(r'(?im)^Attribute VB_Name\s*=\s*"([^"]*)"', text)
        name = found.group(1) if found else path.stem
        valid_name(name)
        self.synchronize()
        names = {entry.name.casefold() for entry in self.entries}
        prefix = name
        counter = 1
        while name.casefold() in names:
            name = f'{prefix}{counter}'
            counter += 1
        entry = VBComponent(self.project, name, kind, '')
        self.entries.append(entry)
        entry.code_module.load_file_text(text)
        entry.pending = True
        self.project.book.saved = False
        return entry


class VBComponent(VBAObject):
    vba_type_name = 'VBComponent'
    vba_library = 'vbide'

    def __init__(self, project: VBProject, name: str, kind: str, source: str) -> None:
        self.project = project
        self.name = name
        self.kind = kind
        self.source = source
        self.pending = False
        self.compile_invalidated = False
        self.invalidated = False
        self.code_module = CodeModule(self)

    def persisted_source(self) -> str:
        if self.code_module.editor_lines is None:
            return self.source
        header = editor_source_parts(self.source)[0]
        lines = self.code_module.editor_lines
        return header + '\r\n'.join(lines) + ('\r\n' if lines else '')

    @member
    def Name(self) -> object:
        return self.name

    @setter('Name')
    def set_name(self, value: object) -> None:
        name = to_text(value)
        valid_name(name)
        if name.casefold() == self.name.casefold():
            return
        if name.casefold() in {'excel', 'vba', 'stdole', 'office'} or any(entry.name.casefold() == name.casefold() for entry in self.project.components.entries):
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
        self.pending = True
        self.compile_invalidated = True
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

    @method
    def Export(self, FileName: object) -> object:
        header = split_attribute_header(self.source)[0]
        attributes = {match.group(1).casefold(): match.group(2) for match in re.finditer(r'(?im)^Attribute (\w+)\s*=\s*([^\r\n]*)', header)}
        if self.kind == 'form':
            raise VBAUnsupportedError('VBComponent.Export(UserForm) is not implemented')
        lines = [f'Attribute VB_Name = "{self.name}"']
        if self.kind != 'standard':
            defaults = {'VB_GlobalNameSpace': 'False', 'VB_Creatable': 'False',
                        'VB_PredeclaredId': str(self.kind == 'document'), 'VB_Exposed': str(self.kind == 'document')}
            lines = ['VERSION 1.0 CLASS', 'BEGIN', "  MultiUse = -1  'True", 'END', *lines]
            lines.extend(f'Attribute {name} = {attributes.get(name.casefold(), value)}' for name, value in defaults.items())
        excluded = {'vb_name', 'vb_base', 'vb_templatederived', 'vb_customizable', 'vb_globalnamespace', 'vb_creatable', 'vb_predeclaredid', 'vb_exposed'}
        lines.extend(line for line in header.splitlines() if line.startswith('Attribute ') and line.split()[1].casefold() not in excluded)
        lines.extend(self.code_module.lines())
        text = '\r\n'.join(lines) + '\r\n'
        try:
            Path(to_text(FileName)).write_bytes(encode_mbcs(text, encoding_for_codepage(self.project.code_page)))
        except OSError:
            raise error(50012, 'Unable to export component') from None
        return EMPTY

class CodeModule(VBAObject):
    vba_type_name = 'CodeModule'
    vba_library = 'vbide'

    def __init__(self, component: VBComponent) -> None:
        self.component = component
        self.editor_lines: list[str] | None = None

    def lines(self) -> list[str]:
        return list(self.editor_lines) if self.editor_lines is not None else body_lines(self.component.source)

    def edit(self, lines: list[str], *, invalidate: bool = True) -> None:
        header = editor_source_parts(self.component.source)[0]
        source = header + '\r\n'.join(lines)
        if source != self.component.source or lines != self.lines():
            self.editor_lines = list(lines)
            self.component.source = source
            self.component.pending = True
            self.component.compile_invalidated |= invalidate
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

    def load_file_text(self, text: str, *, rename: bool = False) -> None:
        attributes = [line for line in text.splitlines() if line.startswith('Attribute ')]
        if rename:
            found = next((re.search(r'^Attribute VB_Name\s*=\s*"([^"]*)"', line) for line in attributes if line.startswith('Attribute VB_Name')), None)
            if found is not None:
                self.component.set_name(found.group(1))
        body = [line if line != 'END' else 'End' for line in text.splitlines() if not line.startswith('Attribute ')]
        header = '\r\n'.join(attributes)
        if header:
            header = re.sub(r'(?m)^(Attribute VB_Name\s*=\s*)"[^"]*"', lambda match: match.group(1) + '"' + self.component.name + '"', header)
            header += '\r\n'
        self.component.source = header + '\r\n'.join(self.lines())
        if body:
            position = int(self.CountOfDeclarationLines())
            lines = self.lines()
            self.edit(lines[:position] + body + lines[position:])

    @method
    def AddFromFile(self, FileName: object) -> object:
        self.load_file_text(read_source_file(Path(to_text(FileName)), self.component.project.code_page), rename=True)
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
        self.edit(lines[:line - 1] + to_text(String).replace('\r\n', '\n').replace('\r', '\n').split('\n') + lines[line:], invalidate=False)
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


def read_source_file(path: Path, code_page: int) -> str:
    try:
        return path.read_bytes().decode(encoding_for_codepage(code_page))
    except FileNotFoundError:
        raise error(53) from None
    except OSError:
        raise error(75) from None
