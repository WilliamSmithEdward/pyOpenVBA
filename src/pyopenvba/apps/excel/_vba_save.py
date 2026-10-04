"""Persist the running Excel project's source through the ordinary VBA writer."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from collections.abc import Sequence

from pyopenvba.apps.excel._model import Workbook
from pyopenvba.excel import ExcelFile
from pyopenvba.interpreter._runtime import ModuleRuntime
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.vba import VBAModuleKind, split_attribute_header


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

    def persist_modules(self, modules: Sequence[ModuleRuntime]) -> bool:
        documents = dict(self._new_documents())
        project = self.vba_project() if self._has_project else self.add_vba_project()
        self._document_names.update(name.casefold() for name in documents)
        existing = {module.name.casefold(): module for module in project.modules}
        changed = self._project_added
        for runtime in modules:
            source = runtime.parsed.source
            kind = VBAModuleKind.standard if runtime.parsed.kind == "standard" else VBAModuleKind.other
            if runtime.parsed.kind == "document":
                header = next((value for name, value in documents.items() if name.casefold() == runtime.name.casefold()), None)
                if header is None:
                    raise ValueError(f"No workbook or sheet document module named {runtime.name!r}")
                if not split_attribute_header(source)[0]:
                    source = header + source
            old = existing.get(runtime.name.casefold())
            if old is None:
                project.add_module(runtime.name, source, kind=kind)
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
    modules = list(interpreter.modules.values())
    if not modules:
        return package
    with MemoryExcelFile(package.serialize()) as host:
        if not host.persist_modules(modules):
            return package
        return OpcFile.parse(host.output_bytes)
