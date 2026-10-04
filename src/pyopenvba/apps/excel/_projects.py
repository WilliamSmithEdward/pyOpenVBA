"""One interpreter per workbook, with nested Excel execution contexts."""
from __future__ import annotations

from collections.abc import Generator
from contextlib import contextmanager

from pyopenvba.apps.excel._bridge import ExcelBridge
from pyopenvba.apps.excel._model import Application, Workbook
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._parse import parse_module
from pyopenvba.interpreter._runtime import Interpreter, MAX_DEPTH, ModuleRuntime, UserClassInstance
from pyopenvba.interpreter._values import EMPTY, error
from pyopenvba.vba import VBAModuleKind, VBAProject


class ExcelInterpreter(Interpreter):
    """Keep VBA globals local while host calls and event handlers nest."""

    def __init__(self, application: Application, workbook: Workbook | None = None) -> None:
        self.bridge = ExcelBridge(application, workbook)
        super().__init__(self.bridge)

    @contextmanager
    def execution_context(self) -> Generator[None, None, None]:
        application = self.bridge.application
        if len(application.executing_projects) >= MAX_DEPTH:
            raise error(28)
        application.executing_projects.append(self)
        try:
            yield
        finally:
            application.executing_projects.pop()

    def initialise(self) -> None:
        with self.execution_context():
            from pyopenvba.apps.excel._vbide import VBProject

            book = self.bridge.workbook
            if book is not None and isinstance(book.vbide_project, VBProject):
                book.vbide_project.compile_pending()
            super().initialise()

    def call(self, procedure: A.Procedure, module: ModuleRuntime, args: list[object],
             named: dict[str, object], *, me: object = None, reset_error: bool = True) -> object:
        with self.execution_context():
            return super().call(procedure, module, args, named, me=me, reset_error=reset_error)

    def add_module(self, source: str, *, name: str = "", kind: str = "standard") -> ModuleRuntime:
        from pyopenvba.apps.excel._documents import materialize

        if self.bridge.workbook is not None:
            materialize(self.bridge.workbook)
        parsed = parse_module(source, name=name, kind=kind)
        if not parsed.name:
            parsed.name = "ThisWorkbook" if kind == "document" else name or f"Module{len(self.modules) + 1}"
        key = parsed.name.casefold()
        previous = self.modules.get(key)
        owner = self.bridge.global_object(parsed.name.casefold()) if kind == "document" else None
        if kind == "document":
            if not isinstance(owner, VBAObject) or owner is self.bridge.application:
                raise ValueError(f"no sheet or workbook has the code name {parsed.name!r}")
        runtime = ModuleRuntime(parsed, self)
        self.modules[key] = runtime
        try:
            if kind == "document":
                assert isinstance(owner, VBAObject)
                self.bind_document(runtime.name, owner)
        except Exception:
            if previous is None:
                del self.modules[key]
            else:
                self.modules[key] = previous
            raise
        book = self.bridge.workbook
        if book is not None:
            book.pending_document_sources.pop(runtime.name, None)
            from pyopenvba.apps.excel._vbide import VBProject

            if isinstance(book.vbide_project, VBProject):
                for entry in book.vbide_project.components.entries:
                    if entry.name.casefold() == runtime.name.casefold():
                        if entry.source != runtime.parsed.source:
                            entry.code_module.editor_lines = None
                        entry.source = runtime.parsed.source
                        entry.pending = False
        if book is not None and (previous is None or previous.parsed.source != runtime.parsed.source):
            book.saved = False
        return runtime

    def load_project(self, project: VBAProject) -> list[str]:
        from pyopenvba.apps.excel._documents import materialize

        if self.bridge.workbook is not None:
            materialize(self.bridge.workbook)
        staged: list[ModuleRuntime] = []
        book = self.bridge.workbook
        was_saved = book.saved if book is not None else True
        for module in project.modules:
            kind = "standard" if module.kind is VBAModuleKind.standard else "class"
            owner = self.bridge.global_object(module.name.lower())
            if kind == "class" and isinstance(owner, VBAObject) and owner is not self.bridge.application:
                kind = "document"
            staged.append(ModuleRuntime(parse_module(module.source, name=module.name, kind=kind), self))
        # Parse every module first so a syntax error cannot silently leave a
        # half-imported project. Install all names before binding documents
        # so their declarations can refer to classes that follow them.
        previous = dict(self.modules)
        previous_documents: list[tuple[VBAObject, UserClassInstance | None]] = []
        try:
            self.modules.update((runtime.name.casefold(), runtime) for runtime in staged)
            for runtime in staged:
                if runtime.parsed.kind == "document":
                    owner = self.bridge.global_object(runtime.name.lower())
                    assert isinstance(owner, VBAObject)
                    previous_documents.append((owner, owner.vba_document))
                    self.bind_document(runtime.name, owner)
        except Exception:
            self.modules.clear()
            self.modules.update(previous)
            for owner, document in previous_documents:
                owner.vba_document = document
            raise
        finally:
            if book is not None:
                book.saved = was_saved
        if book is not None:
            for runtime in staged:
                book.pending_document_sources.pop(runtime.name, None)
                from pyopenvba.apps.excel._vbide import VBProject

                if isinstance(book.vbide_project, VBProject):
                    for entry in book.vbide_project.components.entries:
                        if entry.name.casefold() == runtime.name.casefold():
                            if entry.source != runtime.parsed.source:
                                entry.code_module.editor_lines = None
                            entry.source = runtime.parsed.source
                            entry.pending = False
        return [runtime.name for runtime in staged]


def attach_project(book: Workbook) -> ExcelInterpreter | None:
    if isinstance(book.project_runtime, ExcelInterpreter):
        return book.project_runtime
    root = book.application.interpreter
    if root is None:
        return None
    if isinstance(root, ExcelInterpreter) and root.bridge.workbook is None:
        project = root
        project.bridge.workbook = book
    else:
        project = ExcelInterpreter(book.application, book)
        project.event_sinks = root.event_sinks
        project.err = root.err
    book.project_runtime = project
    return project


def load_project(book: Workbook) -> list[str]:
    """Import the package's project without disturbing another workbook's globals."""
    from pyopenvba.apps.excel._vba_save import MemoryExcelFile

    project = attach_project(book)
    if project is None or book.package is None or not book.package.has("xl/vbaProject.bin"):
        return []
    with MemoryExcelFile(book.package.serialize()) as host:
        return project.load_project(host.vba_project())


def run_macro(application: Application, macro: str, args: list[object]) -> object:
    """Qualified names choose a workbook; VBA unqualified names use the caller."""
    project: Interpreter | None = None
    if "!" in macro:
        book_name, macro = macro.rsplit("!", 1)
        if book_name.startswith("'") and book_name.endswith("'"):
            book_name = book_name[1:-1].replace("''", "'")
        book = next((book for book in application.workbooks_.books if book.name.casefold() == book_name.casefold()), None)
        if book is not None:
            project = book.project_runtime
    elif application.executing_projects:
        project = application.executing_projects[-1]
    elif application.active_book is not None:
        project = application.active_book.project_runtime
    else:
        project = application.interpreter
    if project is None:
        raise error(1004, f"Cannot run macro {macro!r}")
    project.initialise()
    module_name, _, _ = macro.rpartition(".")
    if module_name and module_name.casefold() not in project.modules:
        raise error(1004, f"Cannot run macro {macro!r}")
    procedure, runtime = project.find_procedure(macro, "")
    if procedure is None or runtime is None or (runtime.is_class and runtime.document is None):
        raise error(1004, f"Cannot run macro {macro!r}")
    result = project.run(macro, args, preserve_error=True)
    # Excel runs a document module's procedure and passes its arguments,
    # but Application.Run discards its return value (native probes).
    return EMPTY if runtime.document is not None else result
