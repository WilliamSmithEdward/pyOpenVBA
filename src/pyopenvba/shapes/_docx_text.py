"""A document's text, apart from its shapes.

Word's text is ``w:p`` paragraphs of ``w:r`` runs holding ``w:t``.  A
shape sits in a run too, so reading a paragraph's text means leaving
out what belongs to the drawings inside it.

Writing is deliberately narrow.  A document whose text nobody assigned
is never rewritten, so its formatting, styles and everything else
survive a shape edit untouched.  Assigning ``Content.Text`` is Word's
own way of throwing the whole body away, and this does the same: the
body becomes plain paragraphs, and the shapes go with the text, which
is what live Word does with them.
"""

from __future__ import annotations

import re

from pyopenvba._xml import escape, unescape
from pyopenvba.shapes import _drawingml as dml

_PARAGRAPH = re.compile(r"<w:p\b[^>]*?(?:/>|>.*?</w:p>)", re.DOTALL)
_DRAWING = re.compile(
    r"<mc:AlternateContent\b.*?</mc:AlternateContent>|<w:drawing\b.*?</w:drawing>"
    r"|<w:pict\b.*?</w:pict>",
    re.DOTALL,
)
_SECTION = re.compile(r"<w:sectPr\b[^>]*?(?:/>|>.*?</w:sectPr>)", re.DOTALL)

#: A twip is a twentieth of a point, which is how Word writes a margin.
TWIPS_PER_POINT = 20.0

#: What Word starts a document with: an inch all round.
DEFAULT_MARGIN_POINTS = 72.0


def read_paragraphs(xml: str) -> list[str]:
    """The text of each paragraph, with the drawings left out.

    The drawings go first, before the paragraphs are found at all: a
    shape carries paragraphs of its own inside its text box, and they
    are its text, not the document's.
    """
    body = dml.element_span(xml, "w:body")
    inner = xml[body[0] : body[1]] if body else xml
    inner = _DRAWING.sub(_instead_of_drawing, inner)
    out: list[str] = []
    for found in _PARAGRAPH.finditer(inner):
        pieces: list[str] = []
        for piece in re.finditer(
            r"<w:t\b[^>]*?>(.*?)</w:t>|<w:br\b[^>]*?/?>", found.group(0), re.DOTALL
        ):
            matched = piece.group(0)
            pieces.append("\v" if matched.startswith("<w:br") else unescape(piece.group(1) or ""))
        out.append("".join(pieces))
    return out


#: What an in-line shape is in the text: one character, Chr(1).  A
#: floating shape is none at all, measured in live Word.
INLINE_SHAPE_CHARACTER = "\x01"


def _instead_of_drawing(found: re.Match[str]) -> str:
    """What a drawing leaves behind in the text it was taken out of."""
    markup = found.group(0)
    floating = "<wp:anchor" in markup
    return "" if floating else f'<w:t xml:space="preserve">{INLINE_SHAPE_CHARACTER}</w:t>'


def margins(xml: str) -> tuple[float, float]:
    """The left and top margins in points, which Left and Top are from.

    Word's ``AddShape`` places a shape from the page, and then reports
    its position against the column and the paragraph, so the two
    differ by the margins.
    """
    found = re.search(r"<w:pgMar\b([^>]*?)/?>", xml)
    if found is None:
        return DEFAULT_MARGIN_POINTS, DEFAULT_MARGIN_POINTS
    fields = dict(re.findall(r'([\w:]+)="(-?\d+)"', found.group(1)))
    return (
        _points(fields.get("w:left"), DEFAULT_MARGIN_POINTS),
        _points(fields.get("w:top"), DEFAULT_MARGIN_POINTS),
    )


def _points(twips: str | None, fallback: float) -> float:
    try:
        return int(twips) / TWIPS_PER_POINT if twips is not None else fallback
    except ValueError:
        return fallback


def written_text(paragraphs: list[str], original: str) -> str:
    """The document with this text in it and nothing else.

    Assigning the text takes the shapes with it: asked of live Word, a
    document with five floating shapes and one in-line has none of
    either after ``Content.Text`` is assigned.  Only called when a
    macro assigned it; a document nobody rewrote keeps its own body.
    """
    body = dml.element_span(original, "w:body")
    if body is None:
        return original
    inner = original[body[0] : body[1]]
    head = inner.find(">") + 1
    section = _SECTION.search(inner)
    rebuilt = [f"<w:p>{_runs_for(text)}</w:p>" for text in paragraphs]
    if not rebuilt:
        rebuilt.append("<w:p/>")
    tail = section.group(0) if section is not None else ""
    return original[: body[0]] + inner[:head] + "".join(rebuilt) + tail + "</w:body>" + original[body[1] :]


def _runs_for(text: str) -> str:
    """One paragraph's runs, with a vertical tab as a line break."""
    if not text:
        return ""
    out: list[str] = []
    for index, piece in enumerate(text.split("\v")):
        if index:
            out.append("<w:r><w:br/></w:r>")
        if piece:
            out.append(f'<w:r><w:t xml:space="preserve">{escape(piece)}</w:t></w:r>')
    return "".join(out)
