"""
pyOpenVBA
=========
Read and write VBA sections of Office files using pure Python --
no external dependencies.

Supported formats
-----------------
  Excel:      .xlsm, .xlsb, .xlam (ZIP/OOXML), .xls (raw CFB/BIFF8)
  Word:       .docm, .dotm (ZIP/OOXML), .doc (raw CFB)
  PowerPoint: .pptm, .potm (ZIP/OOXML), .ppt (raw CFB)
  Access:     .accdb, .mdb (the Jet 4 / ACE database file)

Public API
----------
    from pyopenvba import (
        ExcelFile, WordFile, PowerPointFile, AccessDatabase, AccessReader,
        pull, push, pull_word, push_word, pull_ppt, push_ppt,
        pull_access, push_access,
    )

    # In-process module edit (Access)
    with AccessDatabase("database.accdb") as db:
        print(db.module_names())
        db.set_module("Module1", new_src)
        db.save("database_modified.accdb")

    # In-process module edit (Excel)
    with ExcelFile("workbook.xlsm") as wb:
        modules = wb.vba_modules()          # dict[name -> source]
        wb.set_module("Module1", new_src)
        wb.save("workbook_modified.xlsm")

    # In-process module edit (Word)
    with WordFile("document.docm") as doc:
        modules = doc.vba_modules()
        doc.set_module("Module1", new_src)
        doc.save("document_modified.docm")

    # In-process module edit (PowerPoint)
    with PowerPointFile("presentation.pptm") as prs:
        modules = prs.vba_modules()
        prs.set_module("Module1", new_src)
        prs.save("presentation_modified.pptm")

    # Disk-based workflow (.bas / .cls files)
    pull("workbook.xlsm", "./vba_src")         # extract modules (Excel)
    push("./vba_src", "workbook.xlsm")         # write edits back in place

    pull_word("document.docm", "./vba_src")    # extract modules (Word)
    push_word("./vba_src", "document.docm")

    pull_ppt("presentation.pptm", "./vba_src") # extract modules (PowerPoint)
    push_ppt("./vba_src", "presentation.pptm")

    pull_access("database.accdb", "./vba_src") # extract modules (Access)
    push_access("./vba_src", "database.accdb")
"""

from pathlib import Path

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access_read import AccessReader
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import (
    CFBError,
    FormParseError,
    PyOpenVBAError,
    UnsupportedFormatError,
    VBAProjectError,
)
from pyopenvba.forms import FormControl, Size, VBAForm
from pyopenvba.powerpoint import PowerPointFile
from pyopenvba.vba import VBAModuleKind
from pyopenvba.vba import synthesize_class_header as synthesize_class_header
from pyopenvba.word import WordFile


def pull(
    workbook: str | Path,
    dest_dir: str | Path,
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """
    Export every VBA module from ``workbook`` into ``dest_dir`` as
    ``.bas`` / ``.cls`` files. Returns the paths written.
    """
    with ExcelFile(workbook) as wb:
        return wb.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


def push(
    src_dir: str | Path,
    workbook: str | Path,
    *,
    out: str | Path | None = None,
    encoding: str = "utf-8",
    strict: bool = False,
) -> list[str]:
    """
    Update VBA modules in ``workbook`` from ``.bas`` / ``.cls`` files
    in ``src_dir`` and save. Saves in place unless ``out`` is given.
    Returns the list of updated module names.
    """
    with ExcelFile(workbook) as wb:
        updated = wb.push_modules(src_dir, encoding=encoding, strict=strict)
        wb.save(out)
    return updated


def pull_word(
    document: str | Path,
    dest_dir: str | Path,
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """
    Export every VBA module from a Word ``document`` into ``dest_dir`` as
    ``.bas`` / ``.cls`` files. Returns the paths written.
    """
    with WordFile(document) as doc:
        return doc.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


def push_word(
    src_dir: str | Path,
    document: str | Path,
    *,
    out: str | Path | None = None,
    encoding: str = "utf-8",
    strict: bool = False,
) -> list[str]:
    """
    Update VBA modules in a Word ``document`` from ``.bas`` / ``.cls`` files
    in ``src_dir`` and save. Saves in place unless ``out`` is given.
    Returns the list of updated module names.
    """
    with WordFile(document) as doc:
        updated = doc.push_modules(src_dir, encoding=encoding, strict=strict)
        doc.save(out)
    return updated


def pull_ppt(
    presentation: str | Path,
    dest_dir: str | Path,
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """
    Export every VBA module from a PowerPoint ``presentation`` into
    ``dest_dir`` as ``.bas`` / ``.cls`` files. Returns the paths written.
    """
    with PowerPointFile(presentation) as prs:
        return prs.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


def push_ppt(
    src_dir: str | Path,
    presentation: str | Path,
    *,
    out: str | Path | None = None,
    encoding: str = "utf-8",
    strict: bool = False,
) -> list[str]:
    """
    Update VBA modules in a PowerPoint ``presentation`` from ``.bas`` /
    ``.cls`` files in ``src_dir`` and save. Saves in place unless ``out``
    is given. Returns the list of updated module names.
    """
    with PowerPointFile(presentation) as prs:
        updated = prs.push_modules(src_dir, encoding=encoding, strict=strict)
        prs.save(out)
    return updated


def pull_access(
    database: str | Path,
    dest_dir: str | Path,
    *,
    encoding: str = "utf-8",
    overwrite: bool = True,
) -> list[Path]:
    """
    Export every VBA module from an Access ``database`` (.accdb) into
    ``dest_dir`` as ``.bas`` / ``.cls`` files. Returns the paths
    written. Mirrors :func:`pull` / :func:`pull_word` / :func:`pull_ppt`.
    """
    db = AccessReader(database)
    return db.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


def push_access(
    src_dir: str | Path,
    database: str | Path,
    *,
    out: str | Path | None = None,
    encoding: str = "utf-8",
    strict: bool = False,
) -> list[str]:
    """
    Update VBA modules in an Access ``database`` from ``.bas`` / ``.cls``
    files in ``src_dir`` and save. Saves in place unless ``out`` is given.
    Returns the list of updated module names. Mirrors :func:`push`.
    """
    with AccessDatabase(database) as db:
        updated = db.push_modules(src_dir, encoding=encoding, strict=strict)
        db.save(out)
    return updated


__all__ = [
    "AccessDatabase",
    "AccessReader",
    "CFBError",
    "ColumnSpec",
    "ExcelFile",
    "FormControl",
    "FormParseError",
    "IndexSpec",
    "PowerPointFile",
    "PyOpenVBAError",
    "Size",
    "UnsupportedFormatError",
    "VBAForm",
    "VBAModuleKind",
    "VBAProjectError",
    "WordFile",
    "pull",
    "pull_access",
    "pull_ppt",
    "pull_word",
    "push",
    "push_access",
    "push_ppt",
    "push_word",
]

__version__ = "4.0.0"
