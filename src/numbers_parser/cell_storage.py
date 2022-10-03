import math
import re

from collections import OrderedDict
from enum import Enum
from fractions import Fraction
from datetime import timedelta, datetime
from struct import unpack

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


class CustomFormatType(Enum):
    X1 = 1
    X2 = 256
    X3 = 257
    X4 = 258
    X5 = 260
    X6 = 261
    X7 = 262
    X8 = 268
    X9 = 269
    NUMBER = 270
    TEXT = 271
    DATE = 272
    CURRENCY = 274

    def __eq__(self, other):
        if type(other) == int:
            return self.value == other
        else:
            return self.value == other.value


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
        ("DDD", lambda x: str(x.timetuple().tm_yday).zfill(3)),
        ("DD", lambda x: str(x.timetuple().tm_yday).zfill(2)),
        ("D", lambda x: str(x.timetuple().tm_yday).zfill(1)),
        ("HH", "%H"),
        ("H", "%-H"),
        ("hh", "%I"),
        ("h", "%-I"),
        ("k", lambda x: str(x.timetuple().tm_hour).replace("0", "24")),
        ("kk", lambda x: str(x.timetuple().tm_hour).replace("0", "24").zfill(2)),
        ("K", lambda x: str(x.timetuple().tm_hour % 12)),
        ("KK", lambda x: str(x.timetuple().tm_hour % 12).zfill(2)),
        ("mm", lambda x: str(x.timetuple().tm_min).zfill(2)),
        ("m", lambda x: str(x.timetuple().tm_min)),
        ("ss", "%S"),
        ("s", lambda x: str(x.timetuple().tm_sec)),
        ("W", lambda x: week_of_month(x)),
        ("ww", "%W"),
        ("G", "AD"),  # TODO: support BC
        ("F", lambda x: days_occurred_in_month(x)),
        ("S", lambda x: x.strftime("%f")[0]),
        ("SS", lambda x: x.strftime("%f")[0:2]),
        ("SSS", lambda x: x.strftime("%f")[0:3]),
        ("SSSS", lambda x: x.strftime("%f")[0:4]),
        ("SSSSS", lambda x: x.strftime("%f")[0:5]),
    ]
)

CELL_STORAGE_MAP = OrderedDict(
    [
        (0x1, {"attr": "d128", "size": 16}),
        (0x2, {"attr": "double", "size": 8}),
        (0x4, {"attr": "seconds", "size": 8}),
        (0x8, {"attr": "string_id"}),
        (0x10, {"attr": "rich_id"}),
        (0x20, {"attr": "cell_style_id"}),
        (0x40, {"attr": "text_style_id"}),
        (0x80, {"attr": "cond_style_id"}),
        (0x100, {"attr": "cond_rule_style_id"}),
        (0x200, {"attr": "formula_id"}),
        (0x400, {"attr": "control_id"}),
        (0x800, {"attr": "formula_error_id"}),
        (0x1000, {"attr": "suggest_id"}),
        (0x2000, {"attr": "num_format_id"}),
        (0x4000, {"attr": "currency_format_id"}),
        (0x8000, {"attr": "date_format_id"}),
        (0x10000, {"attr": "duration_format_id"}),
        (0x20000, {"attr": "text_format_id"}),
        (0x40000, {"attr": "bool_format_id"}),
        (0x80000, {"attr": "comment_id"}),
        (0x100000, {"attr": "import_warning_id"}),
    ]
)


class CellStorage:
    def __init__(self, model: object, table_id: int, buffer, row_num, col_num):
        version = buffer[0]
        if version != 5:  # pragma: no cover
            raise UnsupportedError(f"Cell storage version {version} is unsupported")

        self._buffer = buffer
        self._model = model
        self._table_id = table_id
        self.row_num = row_num
        self.col_num = col_num

        offset = 12
        flags = unpack("<i", buffer[8:12])[0]
        for mask, field in CELL_STORAGE_MAP.items():
            if flags & mask:
                size = field.get("size", 4)
                if size == 16:
                    value = unpack_decimal128(self._buffer[offset : offset + 16])
                elif size == 8:
                    value = unpack("<d", self._buffer[offset : offset + 8])[0]
                else:
                    value = unpack("<i", self._buffer[offset : offset + 4])[0]
                setattr(self, field["attr"], value)
                offset += size
            else:
                setattr(self, field["attr"], None)

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
            self.value = EPOCH + timedelta(seconds=self.seconds)
            self.datetime = self.value
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

    @property
    def formatted(self):
        if self.duration_format_id is not None and self.double is not None:
            return self.duration_format()
        elif self.date_format_id is not None and self.datetime is not None:
            return self.date_format()
        elif self.num_format_id is not None:
            return self.number_format()
        else:
            return None

    def number_format(self) -> str:
        if self.currency_format_id is not None:
            format = self._model.table_format(self._table_id, self.currency_format_id)
        else:
            format = self._model.table_format(self._table_id, self.num_format_id)
        if format.HasField("custom_uid"):
            format_uuid = uuid(format.custom_uid)
            format_map = self._model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            if custom_format.requires_fraction_replacement:
                accuracy = custom_format.fraction_accuracy
                if accuracy & 0xFF000000:
                    num_digits = 0x100000000 - accuracy
                    formatted_value = float_to_n_digit_fraction(self.d128, num_digits)
                else:
                    formatted_value = float_to_fraction(self.d128, accuracy)
            elif (
                custom_format.format_type == CustomFormatType.NUMBER
                or custom_format.format_type == CustomFormatType.CURRENCY
            ):
                formatted_value = decode_number_format(custom_format, self.d128)
            else:
                raise UnsupportedError(
                    f"Unexpected custom format type {custom_format.format_type}"
                )
        else:
            formatted_value = str(self.d128)
        return formatted_value

    def date_format(self) -> str:
        format = self._model.table_format(self._table_id, self.date_format_id)
        if format.HasField("custom_uid"):
            format_uuid = uuid(format.custom_uid)
            format_map = self._model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            custom_format_string = custom_format.custom_format_string
            if custom_format.format_type == CustomFormatType.DATE:
                formatted_value = decode_date_format(
                    custom_format_string, self.datetime
                )
            else:
                raise UnsupportedError(
                    f"Unexpected custom format type {custom_format.format_type}"
                )
        else:
            formatted_value = decode_date_format(format.date_time_format, self.datetime)
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


def decode_date_format_field(field: str, value: datetime) -> str:
    if field in DATETIME_FIELD_MAP:
        s = DATETIME_FIELD_MAP[field]
        if callable(s):
            return s(value)
        else:
            return value.strftime(s)
    else:  # pragma: no cover
        raise UnsupportedError(f"Unsupported field code '{field}'")


def decode_date_format(format, value):
    """Parse a custom date format string and return a formatted datetime value"""
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
                    result += decode_date_format_field(field, value)
                    in_field = False
                index += 1
        elif in_string:
            result += current_char
            index += 1
        elif not current_char.isalpha():
            if in_field:
                result += decode_date_format_field(field, value)
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
        result += decode_date_format_field(field, value)

    return result


def decode_number_format(format, value):  # noqa: C901
    """Parse a custom date format string and return a formatted number value"""
    custom_format_string = format.custom_format_string
    value *= format.scale_factor
    if "%" in custom_format_string and format.scale_factor == 1.0:
        # Per cent scale has 100x but % does not
        value *= 100.0

    if format.currency_code != "":
        # Replace currency code with symbol and no-break space
        custom_format_string = custom_format_string.replace(
            "\u00a4", format.currency_code + "\u00a0"
        )

    if (match := re.search(r"([#0.,]+(E[+]\d+)?)", custom_format_string)) is not None:
        if "." in match.group(1):
            (int_format_string, dec_format_string) = match.group(1).split(".")
        else:
            (int_format_string, dec_format_string) = (match.group(1), "")
        if not format.show_thousands_separator:
            int_format_string = int_format_string.replace(",", "")

        num_decimals = len(dec_format_string)
        num_integers = len(int_format_string)
        integer = int(value) if num_decimals > 0 else round(value)
        decimal = math.modf(value)[0]

        if match.group(2) is not None:
            formatted_value = f"{value:.{num_decimals - 4}E}"
        elif format.min_integer_width > 0 or format.contains_integer_token:
            num_integers = max([format.min_integer_width, num_integers])
            if format.num_nonspace_integer_digits > 0:
                # Pad integers with zeroes
                if format.show_thousands_separator:
                    formatted_value = f"{integer:0{num_integers},}"
                else:
                    formatted_value = f"{integer:0{num_integers}}"
            else:
                # Pad integers with spaces
                if integer == 0:
                    formatted_value = " " * num_integers
                else:
                    if format.show_thousands_separator:
                        formatted_value = f"{integer:,}".rjust(num_integers)
                    else:
                        formatted_value = str(integer).rjust(num_integers)
        else:
            formatted_value = str(int(integer))

        if match.group(2) is None:
            if format.num_nonspace_decimal_digits > 0:
                # Pad decimal with zeroes
                formatted_value += "." + f"{decimal:,.{num_decimals}f}"[2:]
            elif num_decimals > 0:
                # Pad decimal with spaces
                decimal_str = str(round(decimal, num_decimals))[2:]
                formatted_value += "." + decimal_str.ljust(num_decimals)

        formatted_value = custom_format_string.replace(match.group(1), formatted_value)
    else:
        formatted_value = custom_format_string

    if "'" not in formatted_value:
        return formatted_value

    chars = [*formatted_value]
    index = 0
    in_string = False
    formatted_value = ""
    while index < len(chars):
        current_char = chars[index]
        next_char = chars[index + 1] if index < len(chars) - 1 else None
        if current_char == "'":
            if next_char is None:
                break
            elif chars[index + 1] == "'":
                formatted_value += "'"
                index += 2
            elif in_string:
                in_string = False
                index += 1
            else:
                in_string = True
                index += 1
        else:
            formatted_value += current_char
            index += 1

    return formatted_value


def float_to_fraction(value: float, denominator: int) -> str:
    """Convert a float to the nearest fraction and return as a string"""
    whole = int(value)
    numerator = round(denominator * (value - whole))
    if numerator == 0:
        formatted_value = "0"
    elif whole > 0:
        formatted_value = f"{whole} {numerator}/{denominator}"
    else:
        formatted_value = f"{numerator}/{denominator}"
    return formatted_value


def float_to_n_digit_fraction(value: float, max_digits: int) -> str:
    """Convert a float to a fraction of a maxinum number of digits
    and return as a string"""
    max_denominator = 10**max_digits - 1
    (numerator, denominator) = (
        Fraction.from_float(value).limit_denominator(max_denominator).as_integer_ratio()
    )
    whole = int(value)
    numerator -= whole * denominator
    if numerator == 0:
        formatted_value = "0"
    elif whole == 0:
        formatted_value = f"{numerator}/{denominator}"
    else:
        formatted_value = f"{whole} {numerator}/{denominator}"
    return formatted_value


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
