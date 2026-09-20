"""Shapes on an Office surface: a sheet, a slide, a page.

The vocabulary is in one place because a shape is the same object in
all three hosts, and each host's markup is in its own module beside
it.  What the applications answered about the shapes they made is in
``tests/fixtures/shapes``, measured by ``scripts/measure_shapes.py``,
and every reader and writer here is held to it.
"""

from __future__ import annotations

from pyopenvba.shapes._values import (
    EMU_PER_POINT,
    KINDS,
    MSO_TYPE,
    PRESET_GEOMETRY,
    ControlInfo,
    Shape,
    ShapeKind,
    emu,
    points,
)

__all__ = [
    "ControlInfo",
    "EMU_PER_POINT",
    "KINDS",
    "MSO_TYPE",
    "PRESET_GEOMETRY",
    "Shape",
    "ShapeKind",
    "emu",
    "points",
]
