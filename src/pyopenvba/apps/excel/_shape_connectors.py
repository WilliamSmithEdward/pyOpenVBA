"""What ConnectorFormat connects, as tests/fixtures/shape_connectors.json measured it: each connected end stands on a site of
its shape, and follows it. An elbow or curve reroutes itself by rules not worked out, so only its plain case is built."""
from __future__ import annotations

import math
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error
from pyopenvba.shapes._values import EMU_PER_POINT, ConnectorLink, Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_DIAGONAL = 0.5 * math.cos(math.pi / 4)
_FOUR = ((0.5, 0.0), (0.0, 0.5), (0.5, 1.0), (1.0, 0.5))
#: Where each geometry's connection sites are, across and down the shape, in the order Excel numbers them.
_SITES: dict[str, tuple[tuple[float, float], ...]] = {
    "rect": _FOUR, "roundRect": _FOUR,
    "ellipse": ((0.5, 0.0), (0.5 - _DIAGONAL, 0.5 - _DIAGONAL), (0.0, 0.5), (0.5 - _DIAGONAL, 0.5 + _DIAGONAL), (0.5, 1.0),
                (0.5 + _DIAGONAL, 0.5 + _DIAGONAL), (1.0, 0.5), (0.5 + _DIAGONAL, 0.5 - _DIAGONAL)),
    "triangle": ((0.5, 0.0), (0.25, 0.5), (0.0, 1.0), (0.5, 1.0), (1.0, 1.0), (0.75, 0.5)),
}
_INVALID = (-2147024809, "The specified value is out of range.")


def _sites(target: Shape) -> tuple[tuple[float, float], ...]:
    sites = _SITES.get(target.geometry or "rect")
    if sites is None or target.kind not in ("shape", "textBox"):
        raise VBAUnsupportedError("Connecting to this kind of shape is not implemented")
    return sites


def site_point(target: Shape, site: int) -> tuple[float, float]:
    """Where a site stands on the sheet: within the shape's box, flipped and turned about its centre as the shape is, to the
    file's whole EMU."""
    across, down = _sites(target)[site - 1]
    x = (1 - across if target.flip_h else across) * target.width - target.width / 2
    y = (1 - down if target.flip_v else down) * target.height - target.height / 2
    turn = math.radians(target.rotation)
    return (round((target.left + target.width / 2 + x * math.cos(turn) - y * math.sin(turn)) * EMU_PER_POINT) / EMU_PER_POINT,
            round((target.top + target.height / 2 + x * math.sin(turn) + y * math.cos(turn)) * EMU_PER_POINT) / EMU_PER_POINT)


def _ends(line: Shape) -> tuple[tuple[float, float], tuple[float, float]]:
    near_x, far_x = (line.left + line.width, line.left) if line.flip_h else (line.left, line.left + line.width)
    near_y, far_y = (line.top + line.height, line.top) if line.flip_v else (line.top, line.top + line.height)
    return (near_x, near_y), (far_x, far_y)


def _placed(line: Shape, begin: tuple[float, float], end: tuple[float, float]) -> None:
    line.left, line.top = min(begin[0], end[0]), min(begin[1], end[1])
    line.width, line.height = abs(end[0] - begin[0]), abs(end[1] - begin[1])
    line.flip_h, line.flip_v = end[0] < begin[0], end[1] < begin[1]


def _adjusted(line: Shape, begin: ConnectorLink | None, end: ConnectorLink | None) -> None:
    """An elbow or a curve held at both ends writes its middle out. Held by its beginning alone, with the end behind it, it
    first goes 18 points out from the site; held by its end alone with the beginning behind it, it writes nothing."""
    if line.geometry in ("line", "straightConnector1"):
        return
    if begin is not None and end is not None:
        line.adjustments = (50000,)
    elif begin is not None:
        line.adjustments = (round(-18 / line.width * 100000),)


def _noted(line: Shape) -> None:
    line.link_place = (line.left, line.top, line.flip_h, line.flip_v, line.rotation)


def _guarded(line: Shape, begin: ConnectorLink | None, end: ConnectorLink | None) -> None:
    """An elbow or a curve keeps its plain shape only from a right-hand site towards a left-hand one that lies ahead."""
    if line.geometry in ("line", "straightConnector1"):
        return
    four = all(one is None or one.geometry in ("rect", "roundRect", "") for one in (begin, end))
    plain = four and (begin is None or begin.site == 4) and (end is None or end.site == 2)
    if plain and begin is not None and end is not None and end.point[0] < begin.point[0]:
        plain = False
    near, far = _ends(line)
    if plain and end is None and begin is not None and far[0] >= begin.point[0]:
        plain = False
    if plain and begin is None and end is not None and near[0] > end.point[0]:
        plain = False
    if not plain:
        raise VBAUnsupportedError("Routing an elbow or curved connector this way is not implemented")


def _link(target: Shape, site: int) -> ConnectorLink:
    return ConnectorLink(target.shape_id, site, site_point(target, site), target.geometry)


def connected(sheet: Worksheet, line: Shape, target: Shape, site: int, at_begin: bool) -> None:
    """BeginConnect or EndConnect: the end goes to the site and the other stays where it is."""
    if target is line or target.kind == "group" or (target.kind == "line" and target.geometry != "line"):
        raise error(*_INVALID)
    if target.kind == "line" or not any(one is target for one in sheet.shapes_):
        raise VBAUnsupportedError("Connecting to this kind of shape is not implemented")
    if line.rotation:
        raise VBAUnsupportedError("Connecting a turned connector is not implemented")
    follow(sheet)
    link = _link(target, site)
    begin, end = (link, line.end_link) if at_begin else (line.begin_link, link)
    _guarded(line, begin, end)
    near, far = _ends(line)
    _placed(line, link.point if at_begin else near, far if at_begin else link.point)
    line.begin_link, line.end_link = begin, end
    _adjusted(line, begin, end)
    _noted(line)


def _target(sheet: Worksheet, link: ConnectorLink) -> Shape | None:
    return next((one for one in sheet.shapes_ if one.shape_id == link.target), None)


def follow(sheet: Worksheet) -> None:
    """Bring each connector up to date after a change on the sheet. Moving, flipping or turning the connector itself lets go
    of both its shapes; deleting a shape or changing its type lets go of that end. Where a shape's site has moved, a connector
    held at both ends goes from site to site, and one held at one end moves with it, keeping its size."""
    for line in sheet.shapes_:
        if line.begin_link is None and line.end_link is None:
            continue
        if line.link_place != (line.left, line.top, line.flip_h, line.flip_v, line.rotation):
            line.begin_link, line.end_link, line.link_place = None, None, None
            continue
        current: list[ConnectorLink | None] = []
        moved: tuple[float, float] | None = None
        for link in (line.begin_link, line.end_link):
            target = None if link is None else _target(sheet, link)
            if link is None or target is None or target.geometry != link.geometry:
                current.append(None)
                continue
            fresh = _link(target, link.site)
            if fresh.point != link.point:
                moved = (fresh.point[0] - link.point[0], fresh.point[1] - link.point[1])
            current.append(fresh)
        begin, end = current
        if moved is not None:
            _guarded(line, begin, end)
            if begin is not None and end is not None:
                _placed(line, begin.point, end.point)
            else:
                line.left, line.top = line.left + moved[0], line.top + moved[1]
            _adjusted(line, begin, end)
        line.begin_link, line.end_link = begin, end
        _noted(line)


def rerouted(sheet: Worksheet, line: Shape) -> None:
    """RerouteConnections: the nearest pair of sites, the first such pair if several tie."""
    follow(sheet)
    begin_link, end_link = line.begin_link, line.end_link
    if begin_link is None or end_link is None:
        raise error(*_INVALID)
    begin_target, end_target = _target(sheet, begin_link), _target(sheet, end_link)
    assert begin_target is not None and end_target is not None
    best: tuple[float, int, int] | None = None
    for first in range(1, len(_sites(begin_target)) + 1):
        for second in range(1, len(_sites(end_target)) + 1):
            distance = math.dist(site_point(begin_target, first), site_point(end_target, second))
            if best is None or distance < best[0] - 1e-9:
                best = (distance, first, second)
    assert best is not None
    begin, end = _link(begin_target, best[1]), _link(end_target, best[2])
    _guarded(line, begin, end)
    _placed(line, begin.point, end.point)
    line.begin_link, line.end_link = begin, end
    _adjusted(line, begin, end)
    _noted(line)
