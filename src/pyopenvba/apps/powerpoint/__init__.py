"""A PowerPoint instance in memory, with a VBA interpreter attached.

    >>> from pyopenvba.apps.powerpoint import PowerPointApplication
    >>> app = PowerPointApplication()
    >>> _ = app.add_presentation()
    >>> app.add_module('''
    ... Sub Draw()
    ...     Dim sh As Shape
    ...     Set sh = ActivePresentation.Slides(1).Shapes.AddShape(1, 10, 20, 100, 50)
    ...     sh.TextFrame.TextRange.Text = "Hello"
    ...     sh.ActionSettings(ppMouseClick).Run = "Clicked"
    ... End Sub
    ... ''', name="Module1")
    >>> _ = app.run("Draw")
    >>> app.slide(1).shapes()[0].text
    'Hello'

State comes from a real presentation through
:meth:`PowerPointApplication.open` and goes back out through
:meth:`PowerPointApplication.save`.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Any

from pyopenvba.apps.powerpoint._bridge import PowerPointBridge
from pyopenvba.apps.powerpoint._model import (
    Application,
    Presentation,
    ShapeObject,
    Slide,
)
from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime
from pyopenvba.interpreter._values import EMPTY, MISSING
from pyopenvba.shapes._values import Shape

__all__ = [
    "Application",
    "PowerPointApplication",
    "Presentation",
    "ShapeObject",
    "Slide",
    "SlideView",
]


class SlideView:
    """One slide, from Python rather than from VBA."""

    def __init__(self, slide: Slide) -> None:
        self.slide = slide

    @property
    def name(self) -> str:
        return self.slide.name

    def shapes(self) -> list[Shape]:
        """The shapes on the slide, in the order the file holds them."""
        return list(self.slide.shapes_)

    def shape(self, name: str) -> Shape:
        for one in self.slide.shapes_:
            if one.name.lower() == name.lower():
                return one
        raise KeyError(f"there is no shape called {name!r} on {self.slide.name}")

    def describe(self) -> str:
        lines = [f"{self.slide.name} ({len(self.slide.shapes_)} shapes)"]
        for one in self.slide.shapes_:
            macro = f"  macro={one.macro}" if one.macro else ""
            text = f"  {one.text!r}" if one.text else ""
            lines.append(
                f"  {one.name:20} {one.kind:12} "
                f"{one.left:7.1f},{one.top:7.1f} {one.width:6.1f}x{one.height:6.1f}{macro}{text}"
            )
        return "\n".join(lines)


class PowerPointApplication:
    """PowerPoint, in memory: presentations, slides, shapes, and macros."""

    def __init__(self) -> None:
        self.application = Application()
        self.interpreter = Interpreter(PowerPointBridge(self.application))

    # --- state in -------------------------------------------------------------------

    @classmethod
    def open(cls, path: str | Path, *, with_vba: bool = True) -> PowerPointApplication:
        """Load a presentation, its slides, its shapes and its macros."""
        from pyopenvba.apps.powerpoint._io import load_presentation

        app = cls()
        presentation = load_presentation(app.application, Path(path))
        app.application.presentations_.items.append(presentation)
        app.application.active = presentation
        if with_vba:
            app.load_vba(path)
        return app

    def load_vba(self, path: str | Path) -> list[str]:
        """Add every module of the file's VBA project to the interpreter."""
        from pyopenvba.exceptions import PyOpenVBAError
        from pyopenvba.powerpoint import PowerPointFile
        from pyopenvba.vba import VBAModuleKind

        target = Path(path)
        if target.suffix.lower() not in (".pptm", ".potm", ".ppsm", ".ppt"):
            return []
        names: list[str] = []
        try:
            with PowerPointFile(target) as host:
                project = host.vba_project()
                for module in project.modules:
                    kind = "standard" if module.kind is VBAModuleKind.standard else "class"
                    self.interpreter.add_module(module.source, name=module.name, kind=kind)
                    names.append(module.name)
        except PyOpenVBAError:
            return names
        return names

    def add_presentation(self) -> Presentation:
        """A new presentation with one blank slide."""
        presentation = self.application.presentations_.vba_get("Add")
        assert isinstance(presentation, Presentation)
        return presentation

    def add_module(self, source: str, *, name: str = "", kind: str = "standard") -> ModuleRuntime:
        """Parse VBA and add it to the project this instance runs."""
        return self.interpreter.add_module(source, name=name, kind=kind)

    # --- running --------------------------------------------------------------------

    def run(self, macro: str, *args: object) -> object:
        """Run a macro by name."""
        return _plain(self.interpreter.run(macro, list(args)))

    def evaluate(self, expression: str) -> object:
        """Evaluate one VBA expression against the current state."""
        from pyopenvba.interpreter._parse import parse_expression

        parsed = parse_expression(expression)
        self.interpreter.initialise()
        frame = _bare_frame(self.interpreter)
        return _plain(self.interpreter.evaluate(parsed, frame))

    def freeze_clock(self, when: _dt.datetime) -> None:
        """Pin Now, Date and Time, so two runs can be compared."""
        self.interpreter.freeze(when)

    @property
    def console(self) -> list[str]:
        """Everything Debug.Print wrote."""
        return self.interpreter.console

    # --- state out ------------------------------------------------------------------

    @property
    def presentation(self) -> Presentation:
        presentation = self.application.active
        if presentation is None:
            raise ValueError("no presentation is open")
        return presentation

    def slide(self, index: int) -> SlideView:
        """One slide, counted from one as PowerPoint counts them."""
        slides = self.presentation.slides_
        if not 1 <= index <= len(slides):
            raise IndexError(f"there is no slide {index}")
        return SlideView(slides[index - 1])

    def describe(self) -> str:
        """What is in the presentation, for a console."""
        presentation = self.presentation
        lines = [f"{presentation.name}: {len(presentation.slides_)} slides"]
        for index, slide in enumerate(presentation.slides_, start=1):
            lines.append(f"[{index}] " + SlideView(slide).describe())
        return "\n".join(lines)

    def save(self, path: str | Path | None = None) -> Path:
        """Write the presentation back out, rewriting only what changed."""
        from pyopenvba.apps.powerpoint._io import save_presentation

        presentation = self.presentation
        target = (
            Path(path)
            if path is not None
            else Path(presentation.folder) / presentation.name
        )
        save_presentation(presentation, target)
        presentation.saved = -1
        return target


def _plain(value: object) -> object:
    """A VBA value as an ordinary Python one."""
    if value is EMPTY or value is MISSING:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, float):
        return float(value)
    return value


def _bare_frame(interpreter: Interpreter) -> Any:
    from pyopenvba.interpreter import _ast as A
    from pyopenvba.interpreter._runtime import Frame
    from pyopenvba.interpreter._runtime import ModuleRuntime as _ModuleRuntime

    if interpreter.modules:
        module = next(iter(interpreter.modules.values()))
    else:
        module = _ModuleRuntime(A.Module(name="<immediate>"), interpreter)
    return Frame(procedure=A.Procedure(name="<immediate>"), module=module)
