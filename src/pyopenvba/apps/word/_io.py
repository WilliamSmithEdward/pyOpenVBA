"""Reading a document into the model, and writing it back.

The same rule as everywhere else here: a save puts back only what
changed.  A document whose shapes nobody moved and whose text nobody
assigned is written back byte for byte, styles, headers, numbering and
all.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.exceptions import PyOpenVBAError
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.shapes._docx import read_document, without
from pyopenvba.shapes._docx import written as written_shapes
from pyopenvba.shapes._docx_text import margins, read_paragraphs, written_text

if TYPE_CHECKING:
    from pyopenvba.apps.word._model import Application, Document

#: Where the text of a .docx lives.
BODY = "word/document.xml"


class DocumentFileError(PyOpenVBAError):
    """Raised when a document's parts cannot be read or written."""


def load_document(application: Application, path: Path) -> Document:
    """A document read from a docx, docm, dotm or dotx package."""
    from pyopenvba.apps.word._model import Document

    if not path.exists():
        raise DocumentFileError(f"there is no file at {path}")
    if path.suffix.lower() == ".doc":
        raise DocumentFileError(
            "a .doc document keeps its text in a binary stream that pyOpenVBA does not "
            "read; its VBA project is readable through WordFile"
        )
    package = OpcFile.parse(path.read_bytes())
    if not package.has(BODY):
        raise DocumentFileError(f"{path.name} has no {BODY}")
    document = Document(application, path.name, str(path.parent))
    document.package = package
    document.document_xml = package.read(BODY).decode("utf-8", errors="replace")
    document.paragraphs_ = read_paragraphs(document.document_xml) or [""]
    document.shapes_ = read_document(document.document_xml)
    document.shape_count = len(document.shapes_)
    document.left_margin, document.top_margin = margins(document.document_xml)
    document.saved = True
    document.text_dirty = False
    document.shapes_dirty = False
    return document


def save_document(document: Document, target: Path) -> None:
    """Write the document out, rewriting only what changed."""
    package = document.package
    if package is None:
        raise DocumentFileError(
            "this document was made from nothing, and pyOpenVBA has no template to write "
            "it from; open a .docm and edit that"
        )
    body = document.document_xml
    if document.shapes_dirty:
        for gone in document.deleted_shapes:
            body = without(gone, body)
        document.deleted_shapes.clear()
        body = written_shapes(document.shapes_, body)
    if document.text_dirty:
        body = written_text(document.paragraphs_, body)
    if document.shapes_dirty or document.text_dirty:
        package.write(BODY, body.encode("utf-8"))
        document.document_xml = body
        document.shapes_dirty = False
        document.text_dirty = False
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(package.serialize())
