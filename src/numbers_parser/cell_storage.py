import math
import re

from collections import OrderedDict
from fractions import Fraction
from functools import lru_cache
from pendulum import datetime, duration
from struct import unpack
from typing import Tuple
from warnings import warn

from numbers_parser.constants import (
    EPOCH,
    PACKAGE_ID,
    DECIMAL_PLACES_AUTO,
    CellType,
    CellPadding,
    DurationStyle,
    DurationUnits,
    FormatType,
    SECONDS_IN_HOUR,
    SECONDS_IN_DAY,
    SECONDS_IN_WEEK,
)
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives


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
        ("DDD", lambda x: str(x.day_of_year).zfill(3)),
        ("DD", lambda x: str(x.day_of_year).zfill(2)),
        ("D", lambda x: str(x.day_of_year).zfill(1)),
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
        ("W", lambda x: str(x.week_of_month - 1)),
        ("ww", "%W"),
        ("G", "AD"),  # TODO: support BC
        ("F", lambda x: days_occurred_in_month(x)),
        ("S", lambda x: str(x.microsecond).zfill(6)[0]),
        ("SS", lambda x: str(x.microsecond).zfill(6)[0:2]),
        ("SSS", lambda x: str(x.microsecond).zfill(6)[0:3]),
        ("SSSS", lambda x: str(x.microsecond).zfill(6)[0:4]),
        ("SSSSS", lambda x: str(x.microsecond).zfill(6)[0:5]),
    ]
)

CELL_STORAGE_MAP_V5 = OrderedDict(
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
    def __init__(  # noqa: C901
        self, model: object, table_id: int, buffer, row_num, col_num
    ):
        self.buffer = buffer
        self.model = model
        self.table_id = table_id
        self.row_num = row_num
        self.col_num = col_num

        if buffer is None:
            return

        version = buffer[0]
        if version != 5:  # pragma: no cover
            raise UnsupportedError(f"Cell storage version {version} is unsupported")

        offset = 12
        flags = unpack("<i", buffer[8:12])[0]
        for mask, field in CELL_STORAGE_MAP_V5.items():
            if flags & mask:
                size = field.get("size", 4)
                if size == 16:
                    value = unpack_decimal128(buffer[offset : offset + 16])
                elif size == 8:
                    value = unpack("<d", buffer[offset : offset + 8])[0]
                else:
                    value = unpack("<i", buffer[offset : offset + 4])[0]
                setattr(self, field["attr"], value)
                offset += size
            else:
                setattr(self, field["attr"], None)

        cell_type = buffer[1]
        if cell_type == TSTArchives.genericCellType:
            self.type = CellType.EMPTY
            self.value = None
        elif cell_type == TSTArchives.numberCellType:
            self.value = self.d128
            self.type = CellType.NUMBER
        elif cell_type == TSTArchives.textCellType:
            self.value = self.model.table_string(table_id, self.string_id)
            self.type = CellType.TEXT
        elif cell_type == TSTArchives.dateCellType:
            self.value = EPOCH + duration(seconds=self.seconds)
            self.datetime = self.value
            self.type = CellType.DATE
        elif cell_type == TSTArchives.boolCellType:
            self.value = self.double > 0.0
            self.type = CellType.BOOL
        elif cell_type == TSTArchives.durationCellType:
            self.value = self.double
            self.type = CellType.DURATION
        elif cell_type == TSTArchives.formulaErrorCellType:
            self.value = None
            self.type = CellType.ERROR
        elif cell_type == TSTArchives.automaticCellType:
            self.value = self.model.table_rich_text(self.table_id, self.rich_id)
            self.type = CellType.RICH_TEXT
        elif cell_type == 10:
            self.value = self.d128
            self.type = CellType.NUMBER
        else:  # pragma: no cover
            raise UnsupportedError(f"Cell type ID {cell_type} is not recognised")

    @property
    def formatted(self):
        if self.duration_format_id is not None and self.double is not None:
            return self.duration_format()
        elif self.date_format_id is not None and self.seconds is not None:
            return self.date_format()
        elif (
            self.text_format_id is not None
            or self.num_format_id is not None
            or self.currency_format_id is not None
            or self.bool_format_id is not None
        ):
            return self.custom_format()
        else:
            return str(self.value)

    @property
    @lru_cache(maxsize=None)
    def image_data(self) -> Tuple[bytes, str]:
        """Return the background image data for a cell or None if no image"""
        if self.cell_style_id is None:
            return None
        style = self.model.table_style(self.table_id, self.cell_style_id)
        if not style.cell_properties.cell_fill.HasField("image"):  # pragma: no cover
            return None

        image_id = style.cell_properties.cell_fill.image.imagedata.identifier
        datas = self.model.objects[PACKAGE_ID].datas
        image_file_names = [x.file_name for x in datas if x.identifier == image_id]
        if len(image_file_names) == 0:  # pragma: no cover
            warn(f"No image found with ID {image_id}", UnsupportedWarning)
            return None

        image_path_name = [
            x
            for x in self.model.objects.file_store.keys()
            if x == f"Data/{image_file_names[0]}"
        ][0]
        return (self.model.objects.file_store[image_path_name], image_file_names[0])

    def custom_format(self) -> str:
        if self.text_format_id is not None and self.type == CellType.TEXT:
            format = self.model.table_format(self.table_id, self.text_format_id)
        elif self.currency_format_id is not None:
            format = self.model.table_format(self.table_id, self.currency_format_id)
        elif self.num_format_id is not None:
            format = self.model.table_format(self.table_id, self.num_format_id)
        elif self.bool_format_id is not None:
            format = self.model.table_format(self.table_id, self.bool_format_id)
        else:
            return str(self.value)
        if format.HasField("custom_uid"):
            format_uuid = NumbersUUID(format.custom_uid).hex
            format_map = self.model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            if custom_format.requires_fraction_replacement:
                accuracy = custom_format.fraction_accuracy
                if accuracy & 0xFF000000:
                    num_digits = 0x100000000 - accuracy
                    formatted_value = float_to_n_digit_fraction(self.d128, num_digits)
                else:
                    formatted_value = float_to_fraction(self.d128, accuracy)
            elif custom_format.format_type == FormatType.CUSTOM_TEXT:
                if self.string_id is not None:
                    formatted_value = decode_text_format(
                        custom_format,
                        self.model.table_string(self.table_id, self.string_id),
                    )
                else:
                    return ""
            elif (
                custom_format.format_type == FormatType.CUSTOM_NUMBER
                or custom_format.format_type == FormatType.CUSTOM_CURRENCY
            ):
                formatted_value = decode_number_format(
                    custom_format, self.d128, format_map[format_uuid].name
                )
            else:
                raise UnsupportedError(
                    f"Unexpected custom format type {custom_format.format_type}"
                )
        elif format.format_type == FormatType.DECIMAL:
            return format_decimal(self.d128, format)
        elif format.format_type == FormatType.BOOLEAN:
            return "TRUE" if self.value else "FALSE"
        else:
            formatted_value = str(self.value)
        return formatted_value

    def date_format(self) -> str:
        format = self.model.table_format(self.table_id, self.date_format_id)
        if format.HasField("custom_uid"):
            format_uuid = NumbersUUID(format.custom_uid).hex
            format_map = self.model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            custom_format_string = custom_format.custom_format_string
            if custom_format.format_type == FormatType.CUSTOM_DATE:
                formatted_value = decode_date_format(
                    custom_format_string, self.datetime
                )
            else:  # pragma: no cover
                raise UnsupportedError(
                    f"Unexpected custom format type {custom_format.format_type}"
                )
        else:
            formatted_value = decode_date_format(format.date_time_format, self.datetime)
        return formatted_value

    def duration_format(self) -> str:  # noqa: C901
        format = self.model.table_format(self.table_id, self.duration_format_id)

        duration_style = format.duration_style
        unit_largest = format.duration_unit_largest
        unit_smallest = format.duration_unit_smallest
        if format.use_automatic_duration_units:
            unit_smallest, unit_largest = auto_units(self.double, format)

        d = self.double
        dd = int(self.double)
        dstr = []

        def unit_in_range(largest, smallest, unit_type):
            return largest <= unit_type and smallest >= unit_type

        def pad_digits(d, largest, smallest, unit_type):
            return (largest == unit_type and smallest == unit_type) or d >= 10

        if unit_largest == DurationUnits.WEEK:
            dd = int(d / SECONDS_IN_WEEK)
            if unit_smallest != DurationUnits.WEEK:
                d -= SECONDS_IN_WEEK * dd
            dstr.append(str(dd) + unit_format("week", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.DAY):
            dd = int(d / SECONDS_IN_DAY)
            if unit_smallest > DurationUnits.DAY:
                d -= SECONDS_IN_DAY * dd
            dstr.append(str(dd) + unit_format("day", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.HOUR):
            dd = int(d / SECONDS_IN_HOUR)
            if unit_smallest > DurationUnits.HOUR:
                d -= SECONDS_IN_HOUR * dd
            dstr.append(str(dd) + unit_format("hour", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.MINUTE):
            dd = int(d / 60)
            if unit_smallest > DurationUnits.MINUTE:
                d -= 60 * dd
            if duration_style == DurationStyle.COMPACT:
                pad = pad_digits(dd, unit_smallest, unit_largest, DurationUnits.MINUTE)
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + unit_format("minute", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.SECOND):
            dd = int(d)
            if unit_smallest > DurationUnits.SECOND:
                d -= dd
            if duration_style == DurationStyle.COMPACT:
                pad = pad_digits(dd, unit_smallest, unit_largest, DurationUnits.SECOND)
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + unit_format("second", dd, duration_style))

        if unit_smallest >= DurationUnits.MILLISECOND:
            dd = int(round(1000 * d))
            if duration_style == DurationStyle.COMPACT:
                padding = "0" if dd >= 10 else "00"
                padding = "" if dd >= 100 else padding
                dstr.append(f"{padding}{dd}")
            else:
                dstr.append(
                    str(dd) + unit_format("millisecond", dd, duration_style, "ms")
                )
        duration_str = (":" if duration_style == 0 else " ").join(dstr)
        if duration_style == DurationStyle.COMPACT:
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


def decode_text_format(format, value: str):  # noqa: C901
    """Parse a custom date format string and return a formatted number value"""
    custom_format_string = format.custom_format_string
    return custom_format_string.replace("\ue421", value)


def expand_quotes(value: str) -> str:
    chars = [*value]
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


def decode_number_format(format, value, name):  # noqa: C901
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

    if (
        match := re.search(r"([#0.,]+(E[+]\d+)?)", custom_format_string)
    ) is None:  # pragma: no cover
        warn(
            f"Can't parse format string '{custom_format_string}'; skipping",
            UnsupportedWarning,
        )
        return custom_format_string
    format_spec = match.group(1)
    scientific_spec = match.group(2)

    if format_spec[0] == ".":
        (int_part, dec_part) = ("", format_spec[1:])
    elif "." in custom_format_string:
        (int_part, dec_part) = format_spec.split(".")
    else:
        (int_part, dec_part) = (format_spec, "")

    if scientific_spec is not None:
        # Scientific notation
        formatted_value = f"{value:.{len(dec_part) - 4}E}"
        formatted_value = custom_format_string.replace(format_spec, formatted_value)
        return expand_quotes(formatted_value)

    num_decimals = len(dec_part)
    if num_decimals > 0:
        if dec_part[0] == "#":
            dec_pad = None
        elif format.num_nonspace_decimal_digits > 0:
            dec_pad = CellPadding.ZERO
        else:
            dec_pad = CellPadding.SPACE
    else:
        dec_pad = None
    dec_width = num_decimals

    (integer, decimal) = str(float(value)).split(".")
    if num_decimals > 0:
        integer = int(integer)
        decimal = round(float(f"0.{decimal}"), num_decimals)
    else:
        integer = round(value)
        decimal = float(f"0.{decimal}")

    num_integers = len(int_part.replace(",", ""))
    if not format.show_thousands_separator:
        int_part = int_part.replace(",", "")
    if num_integers > 0:
        if int_part[0] == "#":
            int_pad = None
            int_width = len(int_part)
        elif format.num_nonspace_integer_digits > 0:
            int_pad = CellPadding.ZERO
            if format.show_thousands_separator:
                if integer != 0:
                    num_commas = int(math.floor(math.log10(integer)) / 3)
                else:
                    num_commas = 0
                num_commas = max([num_commas, int((num_integers - 1) / 3)])
                int_width = num_integers + num_commas
            else:
                int_width = num_integers
        else:
            int_pad = CellPadding.SPACE
            int_width = len(int_part)
    else:
        int_pad = None
        int_width = num_integers

    # if num_integers == 0 and dec_pad == CellPadding.SPACE:
    #     dec_pad = CellPadding.ZERO

    # Formatting integer zero:
    #   Blank (padded if needed) if int_pad is SPACE and no decimals
    #   No leading zero if:
    #     int_pad is NONE, dec_pad is SPACE
    #     int_pad is SPACE, dec_pad is SPACE
    #     int_pad is SPACE, dec_pad is ZERO
    #     int_pad is SPACE, dec_pad is NONE if num decimals < decimals length
    if integer == 0 and int_pad == CellPadding.SPACE and num_decimals == 0:
        formatted_value = "".rjust(int_width)
    elif integer == 0 and int_pad is None and dec_pad == CellPadding.SPACE:
        formatted_value = ""
    elif integer == 0 and int_pad == CellPadding.SPACE and dec_pad is not None:
        formatted_value = "".rjust(int_width)
    elif (
        integer == 0
        and int_pad == CellPadding.SPACE
        and dec_pad is None
        and len(str(decimal)) > num_decimals
    ):
        formatted_value = "".rjust(int_width)
    elif int_pad == CellPadding.ZERO:
        if format.show_thousands_separator:
            formatted_value = f"{integer:0{int_width},}"
        else:
            formatted_value = f"{integer:0{int_width}}"
    elif int_pad == CellPadding.SPACE:
        if format.show_thousands_separator:
            formatted_value = f"{integer:,}".rjust(int_width)
        else:
            formatted_value = str(integer).rjust(int_width)
    else:
        if format.show_thousands_separator:
            formatted_value = f"{integer:,}"
        else:
            formatted_value = str(integer)

    if num_decimals:
        if dec_pad == CellPadding.ZERO:
            formatted_value += "." + f"{decimal:,.{dec_width}f}"[2:]
        # elif dec_pad == CellPadding.SPACE and decimal == 0:
        #     formatted_value += ".".ljust(dec_width + 1)
        elif dec_pad == CellPadding.SPACE:
            decimal_str = str(decimal)[2:]
            formatted_value += "." + decimal_str.ljust(dec_width)
        elif decimal or num_integers == 0:
            formatted_value += "." + str(decimal)[2:]

    formatted_value = custom_format_string.replace(format_spec, formatted_value)
    return expand_quotes(formatted_value)


def format_decimal(value, format):
    if value < 0 and format.negative_style == 1:
        accounting_style = False
        value = -value
    elif value < 0 and format.negative_style >= 2:
        accounting_style = True
        value = -value
    else:
        accounting_style = False
    if format.show_thousands_separator:
        thousands = ","
    else:
        thousands = ""
    if value.is_integer() and format.decimal_places >= DECIMAL_PLACES_AUTO:
        formatted_value = f"{int(value):{thousands}}"
    elif format.decimal_places >= DECIMAL_PLACES_AUTO:
        formatted_value = f"{value:{thousands}}"
    else:
        formatted_value = f"{value:{thousands}.{format.decimal_places}f}"
    if accounting_style:
        return f"({formatted_value})"
    else:
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
    if style == DurationStyle.COMPACT:
        return ""
    elif style == DurationStyle.SHORT:
        return f"{abbrev}"
    else:
        return f" {unit}" + plural


def auto_units(cell_value, format):
    unit_largest = format.duration_unit_largest
    unit_smallest = format.duration_unit_smallest

    if cell_value == 0:
        unit_largest = DurationUnits.DAY
        unit_smallest = DurationUnits.DAY
    else:
        if cell_value >= SECONDS_IN_WEEK:
            unit_largest = DurationUnits.WEEK
        elif cell_value >= SECONDS_IN_DAY:
            unit_largest = DurationUnits.DAY
        elif cell_value >= SECONDS_IN_HOUR:
            unit_largest = DurationUnits.HOUR
        elif cell_value >= 60:
            unit_largest = DurationUnits.MINUTE
        elif cell_value >= 1:
            unit_largest = DurationUnits.SECOND
        else:
            unit_largest = DurationUnits.MILLISECOND

        if math.floor(cell_value) != cell_value:
            unit_smallest = DurationUnits.MILLISECOND
        elif cell_value % 60:
            unit_smallest = DurationUnits.SECOND
        elif cell_value % SECONDS_IN_HOUR:
            unit_smallest = DurationUnits.MINUTE
        elif cell_value % SECONDS_IN_DAY:
            unit_smallest = DurationUnits.HOUR
        elif cell_value % SECONDS_IN_WEEK:
            unit_smallest = DurationUnits.DAY
        if unit_smallest < unit_largest:
            unit_smallest = unit_largest

    return unit_smallest, unit_largest
