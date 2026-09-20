"""A Word instance in memory, with a VBA interpreter attached.

    >>> from pyopenvba.apps.word import WordApplication
    >>> app = WordApplication()
    >>> _ = app.add_document()
    >>> app.add_module('''
    ... Sub Draw()
    ...     ActiveDocument.Content.Text = "Hello"
    ...     ActiveDocument.Shapes.AddShape 1, 100, 100, 120, 60
    ... End Sub
    ... ''', name="Module1")
    >>> _ = app.run("Draw")
    >>> app.document.text
    'Hello\\r'

State comes from a real document through :meth:`WordApplication.open`
and goes back out through :meth:`WordApplication.save`.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Any

from pyopenvba.apps.word._bridge import WordBridge
from pyopenvba.apps.word._model import Application, Document, ShapeObject
from pyopenvba.interpreter._runtime import Interpreter, ModuleRuntime
from pyopenvba.interpreter._values import EMPTY, MISSING
from pyopenvba.shapes._values import Shape

__all__ = ["Application", "Document", "ShapeObject", "WordApplication"]


class WordApplication:
    """Word, in memory: documents, their text, their shapes, and macros."""

    def __init__(self) -> None:
        self.application = Application()
        self.interpreter = Interpreter(WordBridge(self.application))

    # --- state in -------------------------------------------------------------------

    @classmethod
    def open(cls, path: str | Path, *, with_vba: bool = True) -> WordApplication:
        """Load a document, its text, its shapes and its macros."""
        from pyopenvba.apps.word._io import load_document

        app = cls()
        document = load_document(app.application, Path(path))
        app.application.documents_.items.append(document)
        app.application.active = document
        if with_vba:
            app.load_vba(path)
        return app

    def load_vba(self, path: str | Path) -> list[str]:
        """Add every module of the file's VBA project to the interpreter."""
        from pyopenvba.exceptions import PyOpenVBAError
        from pyopenvba.vba import VBAModuleKind
        from pyopenvba.word import WordFile

        target = Path(path)
        if target.suffix.lower() not in (".docm", ".dotm", ".doc"):
            return []
        names: list[str] = []
        try:
            with WordFile(target) as host:
                project = host.vba_project()
                for module in project.modules:
                    kind = "standard" if module.kind is VBAModuleKind.standard else "class"
                    self.interpreter.add_module(module.source, name=module.name, kind=kind)
                    names.append(module.name)
        except PyOpenVBAError:
            return names
        return names

    def add_document(self) -> Document:
        """A new empty document, as Word's New does."""
        document = self.application.documents_.vba_get("Add")
        assert isinstance(document, Document)
        return document

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
    def document(self) -> Document:
        document = self.application.active
        if document is None:
            raise ValueError("no document is open")
        return document

    def shapes(self) -> list[Shape]:
        """The floating shapes, in the order Word lists them."""
        return [one for one in self.document.shapes_ if one.placement != "inline"]

    def inline_shapes(self) -> list[Shape]:
        """The shapes that sit in the text."""
        return [one for one in self.document.shapes_ if one.placement == "inline"]

    def describe(self) -> str:
        """What is in the document, for a console."""
        document = self.document
        lines = [
            f"{document.name}: {len(document.paragraphs_)} paragraphs, "
            f"{len(self.shapes())} shapes, {len(self.inline_shapes())} in line"
        ]
        for index, text in enumerate(document.paragraphs_, start=1):
            lines.append(f"  [{index}] {text!r}")
        for one in self.shapes():
            text = f"  {one.text!r}" if one.text else ""
            lines.append(
                f"  {one.name:28} {one.kind:10} "
                f"{one.left:7.1f},{one.top:7.1f} {one.width:6.1f}x{one.height:6.1f}{text}"
            )
        return "\n".join(lines)

    def save(self, path: str | Path | None = None) -> Path:
        """Write the document back out, rewriting only what changed."""
        from pyopenvba.apps.word._io import save_document

        document = self.document
        target = Path(path) if path is not None else Path(document.folder) / document.name
        save_document(document, target)
        document.saved = True
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
