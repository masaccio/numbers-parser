from collections import OrderedDict
from datetime import datetime
from enum import IntEnum
from math import ceil

import enum_tools.documentation

try:
    from importlib.resources import files
# Can't cover exception using modern python
except ImportError:  # pragma: nocover
    from importlib_resources import files

__all__ = [
    "CellPadding",
    "CellType",
    "ControlFormattingType",
    "DurationStyle",
    "DurationUnits",
    "FormatType",
    "FormattingType",
    "FractionAccuracy",
    "NegativeNumberStyle",
    "PaddingType",
]

DEFAULT_DOCUMENT = files("numbers_parser") / "data" / "empty.numbers"

# New document defaults
DEFAULT_COLUMN_COUNT = 8
DEFAULT_COLUMN_WIDTH = 98.0
DEFAULT_PRE_BNC_BYTES = "🤠".encode()  # Yes, really!
DEFAULT_ROW_COUNT = 12
DEFAULT_ROW_HEIGHT = 20.0
DEFAULT_TABLE_OFFSET = 80.0
DEFAULT_TILE_SIZE = 256

# Style defaults
DEFAULT_ALIGNMENT = ("auto", "top")
DEFAULT_BORDER_WIDTH = 0.35
DEFAULT_BORDER_COLOR = (0, 0, 0)
DEFAULT_BORDER_STYLE = "solid"
DEFAULT_FONT = "Helvetica Neue"
DEFAULT_FONT_SIZE = 11.0
DEFAULT_TEXT_INSET = 4.0
DEFAULT_TEXT_WRAP = True
EMPTY_STORAGE_BUFFER = b"\x05\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"

# Formatting values and defaults
DEFAULT_DATETIME_FORMAT = "dd MMM YYY HH:MM"
CHECKBOX_FALSE_VALUE = "☐"
CHECKBOX_TRUE_VALUE = "☑"
STAR_RATING_VALUE = "★"
OPERATOR_PRECEDENCE = {"%": 6, "^": 5, "×": 4, "*": 4, "/": 4, "÷": 4, "+": 3, "-": 3, "&": 2}


# Numbers limits
MAX_TILE_SIZE = 256
MAX_ROW_COUNT = 1000000
MAX_COL_COUNT = 1000
MAX_HEADER_COUNT = 5
MAX_SIGNIFICANT_DIGITS = 15
MAX_BASE = 36

# Root object IDs
DOCUMENT_ID = 1
PACKAGE_ID = 2

# System constants
EPOCH = datetime(2001, 1, 1)  # noqa: DTZ001
SECONDS_IN_HOUR = 60 * 60
SECONDS_IN_DAY = SECONDS_IN_HOUR * 24
SECONDS_IN_WEEK = SECONDS_IN_DAY * 7
DECIMAL128_BIAS = 0x1820

# File format enumerations
DECIMAL_PLACES_AUTO = 253
CURRENCY_CELL_TYPE = 10
CUSTOM_TEXT_PLACEHOLDER = "\ue421"

# Supported versions
SUPPORTED_NUMBERS_VERSIONS = [
    "10.3",
    "11.0",
    "11.1",
    "11.2",
    "12.0",
    "12.1",
    "12.2",
    "13.0",
    "13.1",
    "13.2",
    "14.0",
    "14.1",
    "14.2",
    "14.3",
    "14.4",
]


# Supported date/time directives
def _days_occurred_in_month(value: datetime) -> str:
    """Return how many times the day of the datetime value has fallen in the month."""
    n_days = int((value - value.replace(day=1)).days / 7) + 1
    return str(n_days)


def _day_of_year(value: datetime) -> int:
    """Return the day number in a year for a datetime."""
    return value.timetuple().tm_yday


def _week_of_month(value: datetime) -> int:
    """Return the week number in a month for a datetime."""
    return int(ceil((value.day + value.replace(day=1).weekday()) / 7.0))


DATETIME_FIELD_MAP = OrderedDict(
    [
        ("a", lambda x: x.strftime("%p").lower()),
        ("EEEE", "%A"),
        ("EEE", "%a"),
        ("yyyy", "%Y"),
        ("yy", "%y"),
        ("y", "%Y"),
        ("MMMM", "%B"),
        ("MMM", "%b"),
        ("MM", "%m"),
        ("M", "%-m"),
        ("d", "%-d"),
        ("dd", "%d"),
        ("DDD", lambda x: str(_day_of_year(x)).zfill(3)),
        ("DD", lambda x: str(_day_of_year(x)).zfill(2)),
        ("D", lambda x: str(_day_of_year(x)).zfill(1)),
        ("HH", "%H"),
        ("H", "%-H"),
        ("hh", "%I"),
        ("h", "%-I"),
        ("k", lambda x: str(x.hour).replace("0", "24")),
        ("kk", lambda x: str(x.hour).replace("0", "24").zfill(2)),
        ("K", lambda x: str(x.hour % 12)),
        ("KK", lambda x: str(x.hour % 12).zfill(2)),
        ("mm", lambda x: str(x.minute).zfill(2)),
        ("m", lambda x: str(x.minute)),
        ("ss", "%S"),
        ("s", lambda x: str(x.second)),
        ("W", lambda x: str(_week_of_month(x) - 1)),
        ("ww", "%W"),
        ("G", "AD"),  # TODO: support BC
        ("F", lambda x: _days_occurred_in_month(x)),
        ("S", lambda x: str(x.microsecond).zfill(6)[0]),
        ("SS", lambda x: str(x.microsecond).zfill(6)[0:2]),
        ("SSS", lambda x: str(x.microsecond).zfill(6)[0:3]),
        ("SSSS", lambda x: str(x.microsecond).zfill(6)[0:4]),
        ("SSSSS", lambda x: str(x.microsecond).zfill(6)[0:5]),
    ],
)


class CellType(IntEnum):
    EMPTY = 1
    NUMBER = 2
    TEXT = 3
    DATE = 4
    BOOL = 5
    DURATION = 6
    ERROR = 7
    RICH_TEXT = 8
    CURRENCY = 101
    MERGED = 102


class CellPadding(IntEnum):
    SPACE = 1
    ZERO = 2


class DurationStyle(IntEnum):
    COMPACT = 0
    SHORT = 1
    LONG = 2


class DurationUnits(IntEnum):
    NONE = 0
    WEEK = 1
    DAY = 2
    HOUR = 4
    MINUTE = 8
    SECOND = 16
    MILLISECOND = 32


class FormatType(IntEnum):
    BOOLEAN = 1
    DECIMAL = 256
    CURRENCY = 257
    PERCENT = 258
    SCIENTIFIC = 259
    TEXT = 260
    DATE = 261
    FRACTION = 262
    CHECKBOX = 263
    RATING = 267
    DURATION = 268
    BASE = 269
    CUSTOM_NUMBER = 270
    CUSTOM_TEXT = 271
    CUSTOM_DATE = 272
    CUSTOM_CURRENCY = 274


class FormattingType(IntEnum):
    BASE = 1
    CURRENCY = 2
    DATETIME = 3
    FRACTION = 4
    NUMBER = 5
    PERCENTAGE = 6
    SCIENTIFIC = 7
    TICKBOX = 8
    RATING = 9
    SLIDER = 10
    STEPPER = 11
    POPUP = 12
    TEXT = 13


class ControlFormattingType(IntEnum):
    BASE = 1
    CURRENCY = 2
    FRACTION = 4
    NUMBER = 5
    PERCENTAGE = 6
    SCIENTIFIC = 7


class OwnerKind(IntEnum):
    TABLE_MODEL = 1
    MERGE_OWNER = 5
    HAUNTED_OWNER = 35


FORMATTING_ALLOWED_CELLS = {
    "base": ["NumberCell"],
    "currency": ["NumberCell"],
    "datetime": ["DateCell"],
    "fraction": ["NumberCell"],
    "number": ["NumberCell"],
    "percentage": ["NumberCell"],
    "popup": ["NumberCell", "TextCell"],
    "rating": ["NumberCell"],
    "scientific": ["NumberCell"],
    "slider": ["NumberCell"],
    "stepper": ["NumberCell"],
    "tickbox": ["BoolCell"],
}

FORMATTING_ACTION_CELLS = ["tickbox", "rating", "popup", "slider", "stepper"]


class CustomFormattingType(IntEnum):
    NUMBER = 101
    DATETIME = 102
    TEXT = 103


CUSTOM_FORMATTING_ALLOWED_CELLS = {
    "number": ["NumberCell"],
    "datetime": ["DateCell"],
    "text": ["TextCell"],
}


@enum_tools.documentation.document_enum
class NegativeNumberStyle(IntEnum):
    """
    How negative numbers are formatted.

    This enum is used in cell data formats and cell custom formats using the
    `negative_style` keyword argument.
    """

    MINUS = 0
    """Negative numbers use a simple minus sign."""
    RED = 1
    """Negative numbers are red with no minus sign."""
    PARENTHESES = 2
    """Negative numbers are in parentheses with no minus sign."""
    RED_AND_PARENTHESES = 3
    """Negative numbers are red and in parentheses with no minus sign."""


@enum_tools.documentation.document_enum
class FractionAccuracy(IntEnum):
    """
    How fractions are formatted.

    This enum is used in cell data formats and cell custom formats using the
    `fraction_accuracy` keyword argument.
    """

    THREE = 0xFFFFFFFD
    """Fractions are formatted with up to 3 digits in the denominator."""
    TWO = 0xFFFFFFFE
    """Fractions are formatted with up to 2 digits in the denominator."""
    ONE = 0xFFFFFFFF
    """Fractions are formatted with up to 1 digit in the denominator."""
    HALVES = 2
    """Fractions are formatted to the nearest half."""
    QUARTERS = 4
    """Fractions are formatted to the nearest quarter."""
    EIGTHS = 8
    """Fractions are formatted to the nearest eighth."""
    SIXTEENTHS = 16
    """Fractions are formatted to the nearest sixteenth."""
    TENTHS = 10
    """Fractions are formatted to the nearest tenth."""
    HUNDRETHS = 100
    """Fractions are formatted to the nearest hundredth."""


ALLOWED_FORMATTING_PARAMETERS = {
    FormattingType.BASE: [
        "base",
        "base_places",
        "base_use_minus_sign",
    ],
    FormattingType.CURRENCY: [
        "currency_code",
        "decimal_places",
        "negative_style",
        "show_thousands_separator",
        "use_accounting_style",
    ],
    FormattingType.DATETIME: ["date_time_format"],
    FormattingType.FRACTION: ["fraction_accuracy"],
    FormattingType.NUMBER: [
        "decimal_places",
        "show_thousands_separator",
        "negative_style",
    ],
    FormattingType.PERCENTAGE: [
        "decimal_places",
        "show_thousands_separator",
        "negative_style",
    ],
    FormattingType.SCIENTIFIC: ["decimal_places"],
    FormattingType.POPUP: ["popup_values", "allow_none"],
    FormattingType.RATING: [],
    FormattingType.SLIDER: [
        "control_format",
        "increment",
        "maximum",
        "minimum",
    ],
    FormattingType.STEPPER: [
        "control_format",
        "increment",
        "maximum",
        "minimum",
    ],
    FormattingType.TICKBOX: [],
    FormattingType.TEXT: [],
}

FORMAT_TYPE_MAP = {
    FormattingType.BASE: FormatType.BASE,
    FormattingType.CURRENCY: FormatType.CURRENCY,
    FormattingType.DATETIME: FormatType.DATE,
    FormattingType.FRACTION: FormatType.FRACTION,
    FormattingType.NUMBER: FormatType.DECIMAL,
    FormattingType.PERCENTAGE: FormatType.PERCENT,
    FormattingType.POPUP: FormatType.TEXT,
    FormattingType.RATING: FormatType.RATING,
    FormattingType.SCIENTIFIC: FormatType.SCIENTIFIC,
    FormattingType.SLIDER: FormatType.DECIMAL,
    FormattingType.STEPPER: FormatType.DECIMAL,
    FormattingType.TICKBOX: FormatType.CHECKBOX,
    FormattingType.TEXT: FormatType.TEXT,
}

CUSTOM_FORMAT_TYPE_MAP = {
    CustomFormattingType.NUMBER: FormatType.CUSTOM_NUMBER,
    CustomFormattingType.DATETIME: FormatType.CUSTOM_DATE,
    CustomFormattingType.TEXT: FormatType.CUSTOM_TEXT,
}


class CellInteractionType(IntEnum):
    VALUE_EDITING = 0
    FORMULA_EDITING = 1
    STOCK = 2
    CATEGORY_SUMMARY = 3
    STEPPER = 4
    SLIDER = 5
    RATING = 6
    POPUP = 7
    TOGGLE = 8


CONTROL_CELL_TYPE_MAP = {
    FormattingType.POPUP: CellInteractionType.POPUP,
    FormattingType.SLIDER: CellInteractionType.SLIDER,
    FormattingType.STEPPER: CellInteractionType.STEPPER,
}


@enum_tools.documentation.document_enum
class PaddingType(IntEnum):
    """How integers and decimals are padded in custom number formats."""

    NONE = 0
    """No number padding."""
    ZEROS = 1
    """Pad integers with leading spaces and decimals with trailing spaces."""
    SPACES = 2
    """Pad integers with leading zeroes and decimals with trailing zeroes."""
