"""
Excel file handler.

Supports:
  - .xlsm  (OOXML macro-enabled workbook — ZIP containing xl/vbaProject.bin)
  - .xlsb  (Binary workbook — ZIP containing xl/vbaProject.bin)
  - .xlam  (Macro-enabled add-in — same ZIP container as .xlsm)
  - .xls   (Legacy BIFF8 — the entire file is a CFB)

Usage
-----
    with ExcelFile("book.xlsm") as wb:
        project = wb.vba_project()        # -> VBAProject
        modules = wb.vba_modules()        # -> dict[str, str]
        wb.set_module("Module1", src)
        wb.save("book_out.xlsm")

All shared behavior (read, edit, pull/push, safety-gated save) lives in
:class:`pyopenvba._host.VBAHostFile`.
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba._host import VBAHostFile
from pyopenvba._new_project import WORKBOOK_BASE, document_header, excel_code_name_edits, excel_code_names
from pyopenvba.vba import VBAProject, split_attribute_header

_ZIP_FORMATS = frozenset({".xlsm", ".xlsb", ".xlam"})
_CFB_FORMATS = frozenset({".xls"})
_VBA_ENTRY = "xl/vbaProject.bin"


class ExcelFile(VBAHostFile):
    """
    Open an Excel file and provide access to its VBA project.

    Can be used as a context manager::

        with ExcelFile("book.xlsm") as wb:
            ...
    """

    _zip_formats = _ZIP_FORMATS
    _cfb_formats = _CFB_FORMATS
    _vba_entry = _VBA_ENTRY
    _host_noun = "workbook"
    _application = "Excel"
    _project_storage = "_VBA_PROJECT_CUR"
    _project_formats = frozenset({".xlsm"})
    _main_part = "xl/workbook.xml"

    # ------------------------------------------------------------------
    # A new project, as Excel makes one (see pyopenvba._new_project)
    # ------------------------------------------------------------------

    def _project_template(self) -> bytes:
        from pyopenvba._templates import EMPTY_XLSM_BYTES

        return EMPTY_XLSM_BYTES

    def _new_documents(self) -> list[tuple[str, str | None]]:
        assert self._zip is not None
        book, sheets = excel_code_names(self._zip.namelist(), self._zip.read)
        return [(book, document_header(book, WORKBOOK_BASE)),
                *[(code, document_header(code, base)) for _, code, base in sheets]]

    def _writes_project(self, project: VBAProject) -> bool:
        # Excel writes a project of document modules only once one holds code.
        return any(module.name.casefold() not in self._document_names
                   or split_attribute_header(module.source)[1].strip()
                   for module in project.modules)

    def _project_edits(self) -> dict[str, bytes]:
        assert self._zip is not None
        return excel_code_name_edits(self._zip.namelist(), self._zip.read)

    @classmethod
    def create_new(cls, path: str | Path) -> ExcelFile:
        """
        Create a new macro-enabled workbook at ``path`` containing an empty
        VBA project (``ThisWorkbook``, ``Sheet1``, and a bare ``Module1``)
        and return an open :class:`ExcelFile` for it.

        Supported extensions: ``.xlsm`` (default), ``.xlsb``, and
        ``.xlam`` (Excel add-in).

        The bytes are decoded from a baked-in template captured from a
        freshly Excel-authored file, so the resulting file opens
        cleanly in Excel without any "found a problem" repair prompt.

        ``path`` is overwritten if it already exists.
        """
        target = Path(path)
        suffix = target.suffix.lower()
        if suffix == ".xlsb":
            from pyopenvba._templates import EMPTY_XLSB_BYTES
            template = EMPTY_XLSB_BYTES
        elif suffix == ".xlam":
            from pyopenvba._templates import EMPTY_XLAM_BYTES
            template = EMPTY_XLAM_BYTES
        else:
            from pyopenvba._templates import EMPTY_XLSM_BYTES
            template = EMPTY_XLSM_BYTES
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(template)
        return cls(target)
