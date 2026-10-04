"""Persist the running Excel project's source through the ordinary VBA writer."""

from __future__ import annotations

import io
import zipfile
import re
import struct
from html import escape
from pathlib import Path
from collections.abc import Sequence

from pyopenvba.apps.excel._model import Workbook
from pyopenvba.excel import ExcelFile
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.vba import VBAModuleKind, encoding_for_codepage, parse_project_stream, split_attribute_header


class MemoryExcelFile(ExcelFile):
    """Use the file writer's protection, signature and cache handling in memory."""

    def __init__(self, raw: bytes) -> None:
        self.input_bytes = raw
        self.output_bytes = raw
        super().__init__("memory.xlsm")

    def _open_zip(self) -> None:
        self._container_raw = self.input_bytes
        self._zip = zipfile.ZipFile(io.BytesIO(self.input_bytes), "r")
        self._has_project = self._vba_entry in self._zip.namelist()

    def _write_container(self, dest: str | Path | None, new_cfb_bytes: bytes | None,
                         package_edits: dict[str, bytes | None], *, full_rebuild: bool = False) -> None:
        self.output_bytes = self._serialized_container(new_cfb_bytes, package_edits, full_rebuild=full_rebuild)

    def component_kinds(self) -> dict[str, str]:
        project = self.vba_project()
        information = parse_project_stream(self._get_cfb().get_stream('PROJECT'), code_page=project.code_page)
        return {**{name.casefold(): 'standard' for name in information.standard_modules},
                **{name.casefold(): 'class' for name in information.class_modules},
                **{name.casefold(): 'document' for name, _ in information.document_modules},
                **{name.casefold(): 'form' for name in information.base_classes}}

    def persist_sources(self, sources: Sequence[tuple[str, str, str]], deleted: set[str] | None = None,
                        *, renames: dict[str, str] | None = None, project_name: str = '') -> bool:
        documents = dict(self._new_documents())
        project = self.vba_project() if self._has_project else self.add_vba_project()
        self._document_names.update(name.casefold() for name in documents)
        existing = {module.name.casefold(): module for module in project.modules}
        changed = self._project_added
        for old, new in (renames or {}).items():
            if old.casefold() in existing:
                module = project.rename_module(existing.pop(old.casefold()).name, new)
                existing[new.casefold()] = module
                changed = True
        if project_name and project_name != project.name:
            encoding = encoding_for_codepage(project.code_page)
            old = project.name.encode(encoding)
            new = project_name.encode(encoding)
            record = struct.pack('<HI', 4, len(old)) + old
            prefix = project.dir_raw[:project.dir_modules_offset]
            if prefix.count(record) != 1:
                raise ValueError('Cannot identify the VBA project name record')
            updated = prefix.replace(record, struct.pack('<HI', 4, len(new)) + new, 1)
            project.dir_raw = updated + project.dir_raw[project.dir_modules_offset:]
            project.dir_modules_offset = len(updated)
            cfb = self._get_cfb()
            text = cfb.get_stream('PROJECT').decode(encoding)
            text, count = re.subn(r'(?m)^Name="[^"]*"', lambda _: f'Name="{project_name}"', text, count=1)
            if count != 1:
                raise ValueError('Cannot identify the PROJECT stream name')
            cfb.write_stream('PROJECT', text.encode(encoding))
            project.name = project_name
            project.dir_structure_dirty = True
            self._pending_mutation = True
            changed = True
        for name in deleted or ():
            if name.casefold() in existing:
                project.delete_module(existing.pop(name.casefold()).name)
                changed = True
        for name, header in documents.items():
            if name.casefold() not in existing:
                module = project.add_module(name, header or "", kind=VBAModuleKind.other)
                existing[name.casefold()] = module
                changed = True
        for module_name, source, module_kind in sources:
            kind = VBAModuleKind.standard if module_kind == "standard" else VBAModuleKind.other
            if module_kind == "document":
                header = next((value for name, value in documents.items() if name.casefold() == module_name.casefold()), None)
                if header is None:
                    raise ValueError(f"No workbook or sheet document module named {module_name!r}")
                if not split_attribute_header(source)[0]:
                    source = header + source
            old = existing.get(module_name.casefold())
            if old is None:
                project.add_module(module_name, source, kind=kind)
                changed = True
            else:
                # A bare body preserves the original attributes, including
                # document identity and a UserForm's designer identity.
                comparable = old.source if split_attribute_header(source)[0] else split_attribute_header(old.source)[1]
                if source != comparable:
                    self.set_module(old.name, source)
                    changed = True
        if changed:
            self.save()
        return changed


def persist_project(book: Workbook, package: OpcFile) -> OpcFile:
    """Save the interpreter's owning project; other open books retain theirs.

    Each open workbook's interpreter holds only that project's modules.
    """
    from pyopenvba.apps.excel._io import workbook_format

    interpreter = book.project_runtime
    if interpreter is None or workbook_format(book) == 51:
        return package
    sources = [(runtime.name, runtime.parsed.source, runtime.parsed.kind) for runtime in interpreter.modules.values()]
    sources.extend((name, source, "document") for name, source in book.pending_document_sources.items())
    from pyopenvba.apps.excel._vbide import VBProject

    vbide = book.vbide_project if isinstance(book.vbide_project, VBProject) else None
    if vbide is not None:
        vbide.components.synchronize()
        sources = [(entry.name, entry.source, entry.kind) for entry in vbide.components.entries]
    if not sources and not book.deleted_document_names and vbide is None:
        return package
    from pyopenvba.apps.excel._documents import materialize

    materialize(book)
    _code_names(book, package)
    with MemoryExcelFile(package.serialize()) as host:
        if not host.persist_sources(sources, book.deleted_document_names,
                                    renames=vbide.renames if vbide is not None else None,
                                    project_name=vbide.name if vbide is not None else ''):
            return package
        return OpcFile.parse(host.output_bytes)


def _with_code_name(text: str, property_name: str, root: str, name: str) -> str:
    value = escape(name, quote=True)
    found = re.search(rf'<{property_name}\b[^>]*>', text)
    if found is None:
        return re.sub(rf'(<{root}\b[^>]*>)', rf'\1<{property_name} codeName="{value}"/>', text, count=1)
    tag = found.group()
    if re.search(r'\bcodeName="[^"]*"', tag):
        tag = re.sub(r'\bcodeName="[^"]*"', f'codeName="{value}"', tag)
    else:
        tag = re.sub(r'(/?>)$', rf' codeName="{value}"\1', tag)
    return text[:found.start()] + tag + text[found.end():]


def _code_names(book: Workbook, package: OpcFile) -> None:
    for part, property_name, root, name in [("xl/workbook.xml", "workbookPr", "workbook", book.code_name),
                                            *[(sheet.part_name, "sheetPr", "worksheet", sheet.code_name) for sheet in book.sheets_]]:
        if part and package.has(part):
            text = package.read(part).decode("utf-8")
            changed = _with_code_name(text, property_name, root, name)
            if changed != text:
                package.write(part, changed.encode("utf-8"))
