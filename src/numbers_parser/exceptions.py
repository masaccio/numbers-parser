class NumbersError(Exception):
    """Base class for other exceptions"""

    pass


class UnsupportedError(NumbersError):
    """Raised for unsupported file format features"""

    pass


class NotImplementedError(NumbersError):
    """Raised for missing Protobufs"""

    pass


class FileError(NumbersError):
    """Raised for IO and other OS errors"""

    pass


class FileFormatError(NumbersError):
    """Raised for parsing errors during file load"""

    pass


class FormulaError(NumbersError):
    """ "Raise for formula evaluation errors"""

    pass


class UnsupportedWarning(Warning):
    """Raised for unsupported file format features"""

    pass
