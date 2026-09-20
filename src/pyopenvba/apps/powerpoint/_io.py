"""Reading a presentation into the model, and writing it back.

The same rule as everywhere else here: a save puts back only what
changed.  A slide nobody touched keeps the bytes it arrived with, and
every part the model has no field for -- the masters, the layouts, the
media, the VBA project -- is carried over untouched.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._xml import attributes as _attributes
from pyopenvba.exceptions import PyOpenVBAError
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.shapes._pptx import read_slide, slide_order
from pyopenvba.shapes._pptx import written as written_slide

if TYPE_CHECKING:
    from pyopenvba.apps.powerpoint._model import Application, Presentation


class PresentationFileError(PyOpenVBAError):
    """Raised when a presentation's parts cannot be read or written."""


def load_presentation(application: Application, path: Path) -> Presentation:
    """A presentation read from a pptx, pptm, potm or ppsm package."""
    from pyopenvba.apps.powerpoint._model import Presentation

    if not path.exists():
        raise PresentationFileError(f"there is no file at {path}")
    if path.suffix.lower() == ".ppt":
        raise PresentationFileError(
            "a .ppt presentation keeps its slides in a binary record stream that pyOpenVBA "
            "does not read; its VBA project is readable through PowerPointFile"
        )
    package = OpcFile.parse(path.read_bytes())
    presentation = Presentation(application, path.name, str(path.parent))
    presentation.package = package
    if not package.has("ppt/presentation.xml"):
        raise PresentationFileError(f"{path.name} has no ppt/presentation.xml")
    body = package.read("ppt/presentation.xml").decode("utf-8", errors="replace")
    targets = _relationships(package, "ppt/_rels/presentation.xml.rels")
    for relationship in slide_order(body):
        part = targets.get(relationship, "")
        if not part or not package.has(part):
            continue
        slide = presentation.add_slide()
        slide.part_name = part
        slide.slide_xml = package.read(part).decode("utf-8", errors="replace")
        slide.layout_xml = _layout_of(package, part)
        slide.shapes_ = read_slide(slide.slide_xml, slide.layout_xml)
        slide.dirty = False
    presentation.saved = -1
    return presentation


def _layout_of(package: OpcFile, slide_part: str) -> str:
    """The layout behind a slide, which places its placeholders."""
    folder, _, name = slide_part.rpartition("/")
    targets = _relationships(package, f"{folder}/_rels/{name}.rels", base=folder)
    for part in targets.values():
        if "slideLayout" in part and package.has(part):
            return package.read(part).decode("utf-8", errors="replace")
    return ""


def _relationships(package: OpcFile, rels: str, base: str = "ppt") -> dict[str, str]:
    """Each relationship id against the part it names."""
    if not package.has(rels):
        return {}
    text = package.read(rels).decode("utf-8", errors="replace")
    out: dict[str, str] = {}
    for element in re.findall(r"<Relationship\b[^>]*/>", text):
        fields = _attributes(element)
        target = fields.get("Target", "")
        if not target:
            continue
        out[fields.get("Id", "")] = _resolved(target, base)
    return out


def _resolved(target: str, base: str) -> str:
    """A relationship target as a part name in the package."""
    if target.startswith("/"):
        return target[1:]
    part = f"{base}/{target}"
    while "/../" in part:
        head, _, tail = part.partition("/../")
        part = head.rpartition("/")[0] + "/" + tail
    return part.replace("/./", "/")


def save_presentation(presentation: Presentation, target: Path) -> None:
    """Write the presentation out, rewriting only the slides that changed."""
    package = presentation.package
    if package is None:
        raise PresentationFileError(
            "this presentation was made from nothing, and pyOpenVBA has no template to "
            "write it from; open a .pptm and edit that"
        )
    for slide in presentation.slides_:
        if not slide.dirty or not slide.part_name:
            continue
        package.write(
            slide.part_name,
            written_slide(slide.shapes_, slide.slide_xml).encode("utf-8"),
        )
        slide.slide_xml = package.read(slide.part_name).decode("utf-8", errors="replace")
        slide.dirty = False
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(package.serialize())
