"""
PowerPoint file handler.

Supports:
  - .pptm  (OOXML macro-enabled presentation -- ZIP containing ppt/vbaProject.bin)
  - .potm  (OOXML macro-enabled template -- ZIP containing ppt/vbaProject.bin)
  - .ppt   (Legacy PowerPoint -- a CFB whose 'PowerPoint Document' stream
            embeds the VBA project as a deflated CFB; see
            :mod:`pyopenvba._ppt_container`)

Usage
-----
    with PowerPointFile("presentation.pptm") as prs:
        project = prs.vba_project()       # -> VBAProject
        modules = prs.vba_modules()       # -> dict[str, str]
        prs.set_module("Module1", src)
        prs.save("presentation_out.pptm")

All shared behavior (read, edit, pull/push, safety-gated save) lives in
:class:`pyopenvba._host.VBAHostFile`.
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba._host import VBAHostFile
from pyopenvba.cfb import CFB

_ZIP_FORMATS = frozenset({".pptm", ".potm"})
_CFB_FORMATS = frozenset({".ppt"})
_VBA_ENTRY = "ppt/vbaProject.bin"


class PowerPointFile(VBAHostFile):
    """
    Open a PowerPoint file and provide access to its VBA project.

    Can be used as a context manager::

        with PowerPointFile("presentation.pptm") as prs:
            ...
    """

    _zip_formats = _ZIP_FORMATS
    _cfb_formats = _CFB_FORMATS
    _vba_entry = _VBA_ENTRY
    _host_noun = "presentation"
    _no_vba_hint = (
        "Make sure the presentation has a VBA project "
        "(save as .pptm in PowerPoint)."
    )

    # ------------------------------------------------------------------
    # Legacy container: the VBA project is embedded, not at the root
    # ------------------------------------------------------------------

    def _vba_cfb_bytes(self, container: bytes) -> bytes:
        """Inflate the VBA project CFB out of a binary presentation."""
        from pyopenvba._ppt_container import extract_vba_storage

        return extract_vba_storage(CFB.from_bytes(container))

    def _container_bytes(self, vba_cfb: bytes) -> bytes:
        """Splice a modified project back into the binary presentation."""
        from pyopenvba._ppt_container import replace_vba_storage

        outer = CFB.from_bytes(self._container_raw)
        replace_vba_storage(outer, vba_cfb)
        return outer.to_bytes()

    @classmethod
    def create_new(cls, path: str | Path) -> PowerPointFile:
        """
        Create a new macro-enabled presentation (``.pptm``) at ``path``
        containing an empty VBA project (a bare ``Module1``) and return
        an open :class:`PowerPointFile` for it.

        The bytes are decoded from a baked-in template captured from a
        freshly PowerPoint-authored presentation, so the resulting file
        opens cleanly in PowerPoint without any repair prompt.

        ``path`` is overwritten if it already exists.
        """
        from pyopenvba._templates import EMPTY_PPTM_BYTES

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_PPTM_BYTES)
        return cls(target)
