"""Custom exceptions for pyOpenVBA."""


class PyOpenVBAError(Exception):
    """Base exception for all pyOpenVBA errors."""


class CFBError(PyOpenVBAError):
    """Raised when the Compound File Binary data is malformed or unsupported."""


class VBAProjectError(PyOpenVBAError):
    """Raised when the VBA project structure is invalid."""


class UnsupportedFormatError(PyOpenVBAError):
    """Raised for file formats that pyOpenVBA cannot handle."""


class FormParseError(PyOpenVBAError):
    """Raised when a UserForm's designer streams do not reconcile.

    Deliberately an error rather than a partial control list: a misread
    site array yields plausible-looking controls, and wrong knowledge is
    worse than none.
    """

class PowerQueryError(PyOpenVBAError):
    """Raised when a workbook's Power Query package cannot be read or written.

    The DataMashup blob is a container of containers -- an OPC package, an
    M section document and a metadata document -- and a malformed one is
    refused rather than half-read: a query list built from a broken
    package would look ordinary and be wrong.
    """
