import os

from enum import IntEnum
from pendulum import datetime
from pkg_resources import resource_filename

# New document defaults
DEFAULT_DOCUMENT = resource_filename(__name__, os.path.join("data", "empty.numbers"))
DEFAULT_COLUMN_COUNT = 5
DEFAULT_COLUMN_WIDTH = 98.0
DEFAULT_PRE_BNC_BYTES = "🤠".encode("utf-8")  # Yes, really!
DEFAULT_ROW_COUNT = 10
DEFAULT_ROW_HEIGHT = 20.0
DEFAULT_TABLE_OFFSET = 80.0
DEFAULT_TILE_SIZE = 256


# Numbers limits
MAX_TILE_SIZE = 256
MAX_ROW_COUNT = 1000000
MAX_COL_COUNT = 1000
MAX_HEADER_COUNT = 5

# Root object IDs
DOCUMENT_ID = 1
PACKAGE_ID = 2

# System constants
EPOCH = datetime(2001, 1, 1)
SECONDS_IN_HOUR = 60 * 60
SECONDS_IN_DAY = SECONDS_IN_HOUR * 24
SECONDS_IN_WEEK = SECONDS_IN_DAY * 7

# File format enumerations
DECIMAL_PLACES_AUTO = 253


class CellType(IntEnum):
    EMPTY = 1
    NUMBER = 2
    TEXT = 3
    DATE = 4
    BOOL = 5
    DURATION = 6
    ERROR = 7
    RICH_TEXT = 8


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


class HorizJustification(IntEnum):
    LEFT = 0
    RIGHT = 1
    CENTER = 2
    JUSTIFIED = 3
    AUTO = 4


class VertJustification(IntEnum):
    TOP = 0
    MIDDLE = 1
    BOTTOM = 2


class Justification(IntEnum):
    LEFT_TOP = (HorizJustification.LEFT << 4) | VertJustification.TOP
    LEFT_MIDDLE = (HorizJustification.LEFT << 4) | VertJustification.MIDDLE
    LEFT_BOTTOM = (HorizJustification.LEFT << 4) | VertJustification.BOTTOM
    CENTER_TOP = (HorizJustification.CENTER << 4) | VertJustification.TOP
    CENTER_MIDDLE = (HorizJustification.CENTER << 4) | VertJustification.MIDDLE
    CENTER_BOTTOM = (HorizJustification.CENTER << 4) | VertJustification.BOTTOM
    RIGHT_TOP = (HorizJustification.RIGHT << 4) | VertJustification.TOP
    RIGHT_MIDDLE = (HorizJustification.RIGHT << 4) | VertJustification.MIDDLE
    RIGHT_BOTTOM = (HorizJustification.RIGHT << 4) | VertJustification.BOTTOM
    JUSTIFIED_TOP = (HorizJustification.JUSTIFIED << 4) | VertJustification.TOP
    JUSTIFIED_MIDDLE = (HorizJustification.JUSTIFIED << 4) | VertJustification.MIDDLE
    JUSTIFIED_BOTTOM = (HorizJustification.JUSTIFIED << 4) | VertJustification.BOTTOM
    AUTO_TOP = (HorizJustification.AUTO << 4) | VertJustification.TOP
    AUTO_MIDDLE = (HorizJustification.AUTO << 4) | VertJustification.MIDDLE
    AUTO_BOTTOM = (HorizJustification.AUTO << 4) | VertJustification.BOTTOM
