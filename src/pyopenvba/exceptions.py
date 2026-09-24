"""Custom exceptions for pyOpenVBA."""


class PyOpenVBAError(Exception):
    """Base exception for all pyOpenVBA errors."""


class CFBError(PyOpenVBAError):
    """Raised when the Compound File Binary data is malformed or unsupported."""


class VBAProjectError(PyOpenVBAError):
    """Raised when the VBA project structure is invalid."""


class NoVBAProjectError(VBAProjectError):
    """Raised by a read or write that needs a VBA project the file does not have.

    A macro-enabled file saved before its first macro, or a binary file
    that never held one, has no VBA project at all. That is its normal
    shape, not damage: ``has_vba_project()`` says False, listing reads
    answer empty, and only a read or write of the project, a module, a
    form or a reference raises this. The first macro has to be written
    in the host application, which creates the project.
    """


class UnsupportedFormatError(PyOpenVBAError):
    """Raised for file formats that pyOpenVBA cannot handle."""


class FormParseError(PyOpenVBAError):
    """Raised when a UserForm's designer streams do not reconcile.

    Deliberately an error rather than a partial control list: a misread
    site array yields plausible-looking controls, and wrong knowledge is
    worse than none.
    """

class VBAError(PyOpenVBAError):
    """Base for the three ways evaluating VBA can fail.

    They are kept apart because they mean different things to whoever is
    reading the output.  A compile error is wrong VBA, a run-time error
    is VBA that the host refused, and an unsupported error is a gap in
    pyOpenVBA.  Collapsing the third into either of the others would
    report our own limit as the caller's mistake.
    """

    def __init__(self, message: str, *, where: str = "") -> None:
        super().__init__(f"{where}: {message}" if where else message)
        self.message = message
        self.where = where


class VBACompileError(VBAError):
    """Raised for source VBA itself would refuse to compile.

    The VBA IDE reports these before anything runs, and so does this: a
    module is parsed whole when it is loaded, not statement by statement
    as execution reaches it.
    """


class VBARuntimeError(VBAError):
    """Raised for an error VBA would raise while running.

    Carries the same number and description the ``Err`` object would, and
    is the only one of the three that ``On Error`` can trap.
    """

    def __init__(
        self,
        number: int,
        description: str,
        *,
        source: str = "",
        where: str = "",
    ) -> None:
        super().__init__(f"run-time error {number}: {description}", where=where)
        self.number = number
        self.description = description
        self.source = source


class VBAUnsupportedError(VBAError):
    """Raised for VBA this understands but does not implement.

    A statement, function or object member that real VBA would have run
    and pyOpenVBA cannot.  Deliberately not trappable by ``On Error``:
    swallowing it would let a script carry on over a step that never
    happened and report a result computed from a state that never
    existed.
    """


class PowerQueryError(PyOpenVBAError):
    """Raised when a workbook's Power Query package cannot be read or written.

    The DataMashup blob is a container of containers -- an OPC package, an
    M section document and a metadata document -- and a malformed one is
    refused rather than half-read: a query list built from a broken
    package would look ordinary and be wrong.
    """
