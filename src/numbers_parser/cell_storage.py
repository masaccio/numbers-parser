import math
import re

from collections import OrderedDict
from enum import Enum
from struct import unpack
from datetime import timedelta, datetime

from numbers_parser.exceptions import UnsupportedError
from numbers_parser.constants import EPOCH
from numbers_parser.utils import uuid

from numbers_parser.generated import TSTArchives_pb2 as TSTArchives


class CellType(Enum):
    EMPTY_CELL_TYPE = 1
    NUMBER_CELL_TYPE = 2
    TEXT_CELL_TYPE = 3
    DATE_CELL_TYPE = 4
    BOOL_CELL_TYPE = 5
    DURATION_CELL_TYPE = 6
    ERROR_CELL_TYPE = 7
    BULLET_CELL_TYPE = 8


SECONDS_IN_HOUR = 60 * 60
SECONDS_IN_DAY = SECONDS_IN_HOUR * 24
SECONDS_IN_WEEK = SECONDS_IN_DAY * 7

FORMAT_STYLE_NONE = 0
FORMAT_STYLE_SHORT = 1
FORMAT_STYLE_MEDIUM = 2


CUSTOM_FORMAT_1 = 1
CUSTOM_FORMAT_2 = 256
CUSTOM_FORMAT_3 = 257
CUSTOM_FORMAT_4 = 258
CUSTOM_FORMAT_5 = 260
CUSTOM_FORMAT_6 = 261
CUSTOM_FORMAT_7 = 262
CUSTOM_FORMAT_8 = 268
CUSTOM_FORMAT_9 = 269
CUSTOM_FORMAT_10 = 270
CUSTOM_FORMAT_DATE = 272
CUSTOM_FORMAT_12 = 274

DATETIME_TO_STRFTIME = OrderedDict(
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
        ("HH", "%H"),
        ("H", "%-H"),
        ("h", "%-I"),
        ("kk", "%H"),
        ("mm", "%M"),
        ("ss", "%S"),
        ("W", lambda x: week_of_month(x)),
        ("ww", "%W"),
        ("G", "AD"),  # TODO: support BC
        ("F", lambda x: days_occurred_in_month(x)),
        ("SSS", lambda x: x.strftime("%f")[0:3]),
    ]
)


class CellStorage:
    def __init__(self, model=object, table_id=int, buffer: bytes = None):
        version = buffer[0]
        if version != 5:  # pragma: no cover
            raise UnsupportedError(f"Cell storage version {version} is unsupported")

        self._buffer = buffer
        self._model = model
        self._table_id = table_id
        self._flags = unpack("<i", buffer[8:12])[0]
        self.decode_flags()

        cell_type = buffer[1]
        if cell_type == TSTArchives.genericCellType:
            self.type = CellType.EMPTY_CELL_TYPE
            self.value = None
        elif cell_type == TSTArchives.numberCellType:
            self.value = self.d128
            self.type = CellType.NUMBER_CELL_TYPE
        elif cell_type == TSTArchives.textCellType:
            self.value = self._model.table_string(table_id, self.string_id)
            self.type = CellType.TEXT_CELL_TYPE
        elif cell_type == TSTArchives.dateCellType:
            self.value = self.datetime
            self.type = CellType.DATE_CELL_TYPE
        elif cell_type == TSTArchives.boolCellType:
            self.value = self.double > 0.0
            self.type = CellType.BOOL_CELL_TYPE
        elif cell_type == TSTArchives.durationCellType:
            self.value = self.double
            self.type = CellType.DURATION_CELL_TYPE
        elif cell_type == TSTArchives.formulaErrorCellType:
            self.value = None
            self.type = CellType.ERROR_CELL_TYPE
        elif cell_type == TSTArchives.automaticCellType:
            self.bullets = self._model.table_bullets(self._table_id, self.rich_id)
            self.value = None
            self.type = CellType.BULLET_CELL_TYPE
        elif cell_type == 10:
            self.value = self.d128
            self.type = CellType.NUMBER_CELL_TYPE
        else:  # pragma: no cover
            raise UnsupportedError(f"Cell type ID {cell_type} is not recognised")

    def decode_flags(self):
        self._current_offset = 12

        self.d128 = self.pop_buffer(16) if self._flags & 0x1 else None
        self.double = self.pop_buffer(8) if self._flags & 0x2 else None
        self.datetime = (
            EPOCH + timedelta(seconds=self.pop_buffer(8)) if self._flags & 0x4 else None
        )
        self.string_id = self.pop_buffer() if self._flags & 0x8 else None
        self.rich_id = self.pop_buffer() if self._flags & 0x10 else None
        self.cell_style_id = self.pop_buffer() if self._flags & 0x20 else None
        self.text_style_id = self.pop_buffer() if self._flags & 0x40 else None
        self.cond_style_id = self.pop_buffer() if self._flags & 0x80 else None
        self.cond_rule_style_id = self.pop_buffer() if self._flags & 0x100 else None
        self.formula_id = self.pop_buffer() if self._flags & 0x200 else None
        self.control_id = self.pop_buffer() if self._flags & 0x400 else None
        self.formula_error_id = self.pop_buffer() if self._flags & 0x800 else None
        self.suggest_id = self.pop_buffer() if self._flags & 0x1000 else None
        self.num_format_id = self.pop_buffer() if self._flags & 0x2000 else None
        self.currency_format_id = self.pop_buffer() if self._flags & 0x4000 else None
        self.date_format_id = self.pop_buffer() if self._flags & 0x8000 else None
        self.duration_format_id = self.pop_buffer() if self._flags & 0x10000 else None
        self.text_format_id = self.pop_buffer() if self._flags & 0x20000 else None
        self.bool_format_id = self.pop_buffer() if self._flags & 0x40000 else None
        self.comment_id = self.pop_buffer() if self._flags & 0x80000 else None
        self.import_warning_id = self.pop_buffer() if self._flags & 0x100000 else None

    def pop_buffer(self, size: int = 4):
        offset = self._current_offset
        self._current_offset += size
        if size == 16:
            return unpack_decimal128(self._buffer[offset : offset + 16])
        elif size == 8:
            return unpack("<d", self._buffer[offset : offset + 8])[0]
        else:
            return unpack("<i", self._buffer[offset : offset + 4])[0]

    @property
    def formatted(self):
        if self.duration_format_id is not None and self.double is not None:
            return self.duration_format()
        elif self.date_format_id is not None and self.datetime is not None:
            return self.date_format()
        else:
            return None

    def date_format(self) -> str:
        format = self._model.table_format(self._table_id, self.date_format_id)
        if format.HasField("custom_uid"):
            format_uuid = uuid(format.custom_uid)
            format_map = self._model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            format_spec = custom_format.custom_format_string
            if custom_format.format_type == CUSTOM_FORMAT_DATE:
                formatted_value = expand_custom_format(format_spec, self.datetime)
            else:
                formatted_value = ""
        else:
            formatted_value = expand_custom_format(
                format.date_time_format, self.datetime
            )
        return formatted_value

    def duration_format(self) -> str:  # noqa: C901
        format = self._model.table_format(self._table_id, self.duration_format_id)

        duration_style = format.duration_style
        unit_largest = format.duration_unit_largest
        unit_smallest = format.duration_unit_smallest
        if format.use_automatic_duration_units:
            unit_smallest, unit_largest = auto_units(self.double, format)

        d = self.double
        dd = int(self.double)
        dstr = []

        if unit_largest == 1:
            dd = int(d / SECONDS_IN_WEEK)
            if unit_smallest != 1:
                d -= SECONDS_IN_WEEK * dd
            dstr.append(str(dd) + unit_format("week", dd, duration_style))
        if unit_largest <= 2 and unit_smallest >= 2:
            dd = int(d / SECONDS_IN_DAY)
            if unit_smallest > 2:
                d -= SECONDS_IN_DAY * dd
            dstr.append(str(dd) + unit_format("day", dd, duration_style))
        if unit_largest <= 4 and unit_smallest >= 4:
            dd = int(d / SECONDS_IN_HOUR)
            if unit_smallest > 4:
                d -= SECONDS_IN_HOUR * dd
            dstr.append(str(dd) + unit_format("hour", dd, duration_style))
        if unit_largest <= 8 and unit_smallest >= 8:
            dd = int(d / 60)
            if unit_smallest > 8:
                d -= 60 * dd
            if duration_style == FORMAT_STYLE_NONE:
                pad = (unit_largest == 8 and unit_smallest == 8) or dd > 10
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + unit_format("minute", dd, duration_style))
        if unit_largest <= 16 and unit_smallest >= 16:
            dd = int(d)
            if unit_smallest > 16:
                d -= dd
            if duration_style == FORMAT_STYLE_NONE:
                pad = (unit_smallest == 16 and unit_largest == 16) or dd >= 10
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + unit_format("second", dd, duration_style))
        if unit_smallest >= 32:
            dd = int(round(1000 * d))
            if duration_style == FORMAT_STYLE_NONE:
                padding = "0" if dd >= 10 else "00"
                padding = "" if dd >= 100 else padding
                dstr.append(f"{padding}{dd}")
            else:
                dstr.append(
                    str(dd) + unit_format("millisecond", dd, duration_style, "ms")
                )
        duration_str = (":" if duration_style == 0 else " ").join(dstr)
        if duration_style == FORMAT_STYLE_NONE:
            duration_str = re.sub(r":(\d\d\d)$", r".\1", duration_str)
        return duration_str


def unpack_decimal128(buffer: bytearray) -> float:
    exp = (((buffer[15] & 0x7F) << 7) | (buffer[14] >> 1)) - 0x1820
    mantissa = buffer[14] & 1
    for i in range(13, -1, -1):
        mantissa = mantissa * 256 + buffer[i]
    if buffer[15] & 0x80:
        mantissa = -mantissa
    value = mantissa * 10**exp
    return float(value)


def week_of_month(value: datetime) -> str:
    """Return the week of the month for a datetime value"""
    month_week = value.isocalendar().week - value.replace(day=1).isocalendar().week
    return str(month_week)


def days_occurred_in_month(value: datetime) -> str:
    """Return how many times the day of the datetime value has fallen in the month"""
    n_days = int((value - value.replace(day=1)).days / 7) + 1
    return str(n_days)


def replace_format_field(field: str, value: datetime) -> str:
    if field in DATETIME_TO_STRFTIME:
        s = DATETIME_TO_STRFTIME[field]
        if callable(s):
            return s(value)
        else:
            return value.strftime(s)
    else:  # pragma: no cover
        raise UnsupportedError(f"Unsupported field code '{field}'")


def expand_custom_format(format, value):
    chars = [*format]
    index = 0
    in_string = False
    in_field = False
    result = ""
    field = ""
    while index < len(chars):
        current_char = chars[index]
        next_char = chars[index + 1] if index < len(chars) - 1 else None
        if current_char == "'":
            if next_char is None:
                break
            elif chars[index + 1] == "'":
                result += "'"
                index += 2
            elif in_string:
                in_string = False
                index += 1
            else:
                in_string = True
                if in_field:
                    result += replace_format_field(field, value)
                    in_field = False
                index += 1
        elif in_string:
            result += current_char
            index += 1
        elif not current_char.isalpha():
            if in_field:
                result += replace_format_field(field, value)
                in_field = False
            result += current_char
            index += 1
        elif in_field:
            field += current_char
            index += 1
        else:
            in_field = True
            field = current_char
            index += 1
    if in_field:
        result += replace_format_field(field, value)

    return result


def unit_format(unit: str, value: int, style: int, abbrev: str = None):
    plural = "" if value == 1 else "s"
    if abbrev is None:
        abbrev = unit[0]
    if style >= FORMAT_STYLE_MEDIUM:
        return f" {unit}" + plural
    elif style == FORMAT_STYLE_SHORT:
        return f"{abbrev}"
    else:
        return ""


def auto_units(cell_value, format):
    unit_largest = format.duration_unit_largest
    unit_smallest = format.duration_unit_smallest

    if cell_value == 0:
        unit_largest = 2
        unit_smallest = 2
    else:
        if cell_value >= SECONDS_IN_WEEK:
            unit_largest = 1
        elif cell_value >= SECONDS_IN_DAY:
            unit_largest = 2
        elif cell_value >= SECONDS_IN_HOUR:
            unit_largest = 4
        elif cell_value >= 60:
            unit_largest = 8
        elif cell_value >= 1:
            unit_largest = 16
        else:
            unit_largest = 32

        if math.floor(cell_value) != cell_value:
            unit_smallest = 32
        elif cell_value % 60:
            unit_smallest = 16
        elif cell_value % SECONDS_IN_HOUR:
            unit_smallest = 8
        elif cell_value % SECONDS_IN_DAY:
            unit_smallest = 4
        elif cell_value % SECONDS_IN_WEEK:
            unit_smallest = 2
        if unit_smallest < unit_largest:
            unit_smallest = unit_largest

    return unit_smallest, unit_largest
