"""
Word file handler.

Supports:
  - .docm  (OOXML macro-enabled document -- ZIP containing word/vbaProject.bin)
  - .dotm  (OOXML macro-enabled template -- ZIP containing word/vbaProject.bin)
  - .doc   (Legacy Word -- the entire file is a CFB)

Usage
-----
    with WordFile("document.docm") as doc:
        project = doc.vba_project()       # -> VBAProject
        modules = doc.vba_modules()       # -> dict[str, str]
        doc.set_module("Module1", src)
        doc.save("document_out.docm")

All shared behavior (read, edit, pull/push, safety-gated save) lives in
:class:`pyopenvba._host.VBAHostFile`.
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba._host import VBAHostFile

_ZIP_FORMATS = frozenset({".docm", ".dotm"})
_CFB_FORMATS = frozenset({".doc"})
_VBA_ENTRY = "word/vbaProject.bin"


class WordFile(VBAHostFile):
    """
    Open a Word file and provide access to its VBA project.

    Can be used as a context manager::

        with WordFile("document.docm") as doc:
            ...
    """

    _zip_formats = _ZIP_FORMATS
    _cfb_formats = _CFB_FORMATS
    _vba_entry = _VBA_ENTRY
    _host_noun = "document"
    _no_vba_hint = (
        "Make sure the document has a VBA project (save as .docm in Word)."
    )

    @classmethod
    def create_new(cls, path: str | Path) -> WordFile:
        """
        Create a new macro-enabled document (``.docm``) at ``path``
        containing an empty VBA project (``ThisDocument`` and a bare
        ``Module1``) and return an open :class:`WordFile` for it.

        The bytes are decoded from a baked-in template captured from a
        freshly Word-authored document, so the resulting file opens
        cleanly in Word without any repair prompt.

        ``path`` is overwritten if it already exists.
        """
        from pyopenvba._templates import EMPTY_DOCM_BYTES

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_DOCM_BYTES)
        return cls(target)
