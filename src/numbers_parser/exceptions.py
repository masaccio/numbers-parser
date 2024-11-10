class NumbersError(Exception):
    """Base class for other exceptions."""


class UnsupportedError(NumbersError):
    """Raised for unsupported file format features."""


class NotImplementedError(NumbersError):
    """Raised for unsuported Protobufs/Formats."""


class FileError(NumbersError):
    """Raised for IO and other OS errors."""


class FileFormatError(NumbersError):
    """Raised for parsing errors during file load."""


class FormulaError(NumbersError):
    """ "Raise for formula evaluation errors."""


class UnsupportedWarning(Warning):
    """Raised for unsupported file format features."""
