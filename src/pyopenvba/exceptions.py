"""Custom exceptions for pyOpenVBA."""


class PyOpenVBAError(Exception):
    """Base exception for all pyOpenVBA errors."""


class CFBError(PyOpenVBAError):
    """Raised when the Compound File Binary data is malformed or unsupported."""


class VBAProjectError(PyOpenVBAError):
    """Raised when the VBA project structure is invalid."""


class UnsupportedFormatError(PyOpenVBAError):
    """Raised for file formats that pyOpenVBA cannot handle."""
