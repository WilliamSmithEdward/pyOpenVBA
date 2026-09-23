"""The small amount of XML handling the OPC writers share.

Office's parts are edited as text rather than re-serialised from a tree,
so that everything the model has no field for survives a save.  These
three helpers are what that needs: the attributes of one element, and
the two escapes.
"""

from __future__ import annotations

import re

_ATTRIBUTE = re.compile(r'([\w.:-]+)\s*=\s*"([^"]*)"')
_ENTITY = re.compile(r"&(#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z]+);")
_NAMED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def unescape(text: str) -> str:
    """The characters an XML attribute's stored text stands for.

    A sheet named ``A & B`` is stored as ``A &amp; B``, and reading the
    stored form as if it were the name meant that asking for the sheet
    by the name it actually has found nothing.  Numeric references are
    read too, since a writer may use one for any character at all.
    """

    def replace(match: re.Match[str]) -> str:
        body = match.group(1)
        if body[:2] in ("#x", "#X"):
            return chr(int(body[2:], 16))
        if body.startswith("#"):
            return chr(int(body[1:]))
        return _NAMED.get(body, match.group(0))

    return _ENTITY.sub(replace, text)


def escape(text: str) -> str:
    """Text as XML spells it."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def escape_text(text: str) -> str:
    """Text between tags as Excel spells it in a cell's formula and value: a quote left as it is."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def attributes(element: str) -> dict[str, str]:
    """The attributes of one element, in whatever order they were written.

    XML gives attribute order no meaning, and writers differ: Excel opens
    a relationship with ``Id``, openpyxl closes with it.  Matching them in
    a fixed order silently found nothing in the second case, which left a
    sheet looking as though it had no part behind it.

    Values come back as the characters they stand for, not as stored, so
    a caller comparing one against a name a user typed compares like with
    like.

    Only the element's own tag is read.  Handing this a whole ``<row>``
    with its cells inside would otherwise answer with the last cell's
    attributes; :func:`tag_attributes` does that slicing.
    """
    return {key: unescape(value) for key, value in _ATTRIBUTE.findall(element)}


def tag_attributes(element: str) -> dict[str, str]:
    """The attributes of an element's opening tag, ignoring its content."""
    stop = element.find(">")
    return attributes(element if stop < 0 else element[: stop + 1])
