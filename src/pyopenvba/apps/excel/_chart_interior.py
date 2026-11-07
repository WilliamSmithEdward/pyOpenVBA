"""Interior on a chart's area, plot area, legend, title and series: the classic view of the fill that Format also reaches.

Measured in live Excel, on a plain chart and one AddChart2 made (scripts/measure_chart_interior.py, tests/fixtures/chart_interior.json).
An automatic fill reads its colour, or -2 where the element has no automatic fill (a legend, a title, a styled chart's plot area), and
-4105 for its index and pattern. A colour of its own reads that colour, the palette's nearest index, pattern 1 and pattern index 2. A
hidden fill reads -2 and -4142. Pattern xlSolid keeps the colour shown as its own, black where none was, and xlNone hides the fill.
What a hidden fill's pattern reads depends on how it was hidden where the element's automatic fill shows, so that read is refused.
"""

from __future__ import annotations

from typing import cast

from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, setter
from pyopenvba.interpreter._values import VBAInt, to_integer, to_number

_AUTOMATIC, _NONE, _SOLID, _NO_COLOUR = -4105, -4142, 1, -2
#: The elements with no automatic fill, whatever the chart's style.
_NONE_KINDS = ('Legend', 'ChartTitle')


class ChartInterior(Created):
    vba_type_name = 'Interior'

    def __init__(self, element: VBAObject, chart: object) -> None:
        self.element, self.chart = element, chart

    def _fill(self) -> VBAObject:
        state = getattr(self.chart, 'state')
        if state.original and state.style_part:
            # Excel reads a fill a loaded styled chart's style wrote out as the element's own: a series' index reads 46, where a
            # new chart of the style reads -4105 (tests/fixtures/loaded_chart_looks.json).
            raise VBAUnsupportedError('The Interior of an element of an imported styled chart is not implemented')
        return cast(VBAObject, cast(VBAObject, self.element.vba_get('Format')).vba_get('Fill'))

    def _colour(self) -> VBAObject:
        return cast(VBAObject, self._fill().vba_get('ForeColor'))

    def _state(self) -> str:
        """'hidden', 'own' or 'automatic', from what the fill reads and holds."""
        fill = self._fill()
        if int(cast(int, fill.vba_get('Visible'))) == 0:
            return 'hidden'
        look = getattr(fill, 'usable', None) or getattr(fill, 'look')
        held = look()
        return 'own' if held.fill_rgb is not None else 'automatic'

    def _colourless(self) -> bool:
        """An automatic fill that is none: Format reads it visible, at colour 0."""
        return bool(getattr(self._colour(), 'colourless')())

    def _none_kind(self) -> bool:
        from pyopenvba.apps.excel._chart_styled import is_styled

        kind = self.element.vba_type_name
        state = getattr(self.chart, 'state')
        return kind in _NONE_KINDS or (kind == 'PlotArea' and (is_styled(state.chart_style) or state.style_part))

    def _palette(self) -> object:
        return getattr(self.chart, 'sheet').book.stylesheet.colors

    @member
    def Color(self) -> object:
        state = self._state()
        if state == 'hidden' or (state == 'automatic' and self._colourless()):
            return VBAInt(_NO_COLOUR, 'Long')
        return self._colour().vba_get('RGB')

    @setter('Color')
    def _color(self, value: object) -> None:
        self._colour().vba_set('RGB', value)

    @member
    def ColorIndex(self) -> object:
        state = self._state()
        if state == 'hidden':
            return VBAInt(_NONE, 'Long')
        if state == 'automatic':
            return VBAInt(_AUTOMATIC, 'Long')
        bgr = int(cast(int, self._colour().vba_get('RGB')))
        rgb = f'{bgr & 0xFF:02X}{(bgr >> 8) & 0xFF:02X}{(bgr >> 16) & 0xFF:02X}'
        return VBAInt(getattr(self._palette(), 'nearest')(rgb), 'Long')

    @setter('ColorIndex')
    def _color_index(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted == _NONE:
            self._fill().vba_set('Visible', 0)
        elif wanted == _AUTOMATIC:
            if self._state() != 'automatic':
                raise VBAUnsupportedError('Giving a chart element its automatic fill again is not implemented')
        elif 1 <= wanted <= 56:
            rgb = getattr(self._palette(), 'indexed')[wanted + 7]
            self._colour().vba_set('RGB', int(rgb[4:6] + rgb[2:4] + rgb[:2], 16))
        else:
            raise VBAUnsupportedError(f'Interior.ColorIndex = {wanted} on a chart element is not implemented')

    def _patterned(self, own: int) -> object:
        state = self._state()
        if state == 'hidden':
            if not self._none_kind():
                raise VBAUnsupportedError("The pattern of a chart element's hidden fill is not implemented")
            return VBAInt(_NONE, 'Long')
        return VBAInt(own if state == 'own' else _AUTOMATIC, 'Long')

    @member
    def Pattern(self) -> object:
        return self._patterned(_SOLID)

    @setter('Pattern')
    def _pattern(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted == _NONE:
            self._fill().vba_set('Visible', 0)
            return
        if wanted != _SOLID:
            raise VBAUnsupportedError(f'Interior.Pattern = {wanted} on a chart element is not implemented')
        state = self._state()
        if state == 'hidden':
            raise VBAUnsupportedError("A solid pattern on a chart element's hidden fill is not implemented")
        if state == 'automatic':
            shown = 0 if self._colourless() else int(to_number(self._colour().vba_get('RGB')))
            self._colour().vba_set('RGB', shown)

    @member
    def PatternColorIndex(self) -> object:
        return self._patterned(2)

    @member
    def InvertIfNegative(self) -> object:
        self._state()
        return False

    @member
    def Parent(self) -> object:
        return self.element

    @member
    def Application(self) -> object:
        return getattr(self.chart, 'sheet').book.application
