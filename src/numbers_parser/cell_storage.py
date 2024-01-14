import logging
import math
import re
from fractions import Fraction
from struct import unpack
from typing import Tuple, Union
from warnings import warn

import sigfig
from pendulum import datetime, duration

from numbers_parser import __name__ as numbers_parser_name
from numbers_parser.constants import (
    CURRENCY_CELL_TYPE,
    DATETIME_FIELD_MAP,
    DECIMAL_PLACES_AUTO,
    EPOCH,
    MAX_SIGNIFICANT_DIGITS,
    PACKAGE_ID,
    SECONDS_IN_DAY,
    SECONDS_IN_HOUR,
    SECONDS_IN_WEEK,
    CellPadding,
    CellType,
    CustomFormattingType,
    DurationStyle,
    DurationUnits,
    FormattingType,
    FormatType,
)
from numbers_parser.currencies import CURRENCY_SYMBOLS
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.numbers_cache import Cacheable, cache
from numbers_parser.numbers_uuid import NumbersUUID

logger = logging.getLogger(numbers_parser_name)
debug = logger.debug


class CellStorage(Cacheable):
    # 15% performance uplift for using slots
    __slots__ = (
        "buffer",
        "datetime",
        "model",
        "table_id",
        "row_num",
        "col_num",
        "value",
        "type",
        "d128",
        "double",
        "seconds",
        "string_id",
        "rich_id",
        "cell_style_id",
        "text_style_id",
        # "cond_style_id",
        # "cond_rule_style_id",
        "formula_id",
        # "control_id",
        "formula_error_id",
        "suggest_id",
        "num_format_id",
        "currency_format_id",
        "date_format_id",
        "duration_format_id",
        "text_format_id",
        "bool_format_id",
        # "comment_id",
        # "import_warning_id",
        "is_currency",
        "_cache",
    )

    # @profile
    def __init__(  # noqa: PLR0912, PLR0913, PLR0915
        self, model: object, table_id: int, buffer, row_num, col_num
    ):
        self.buffer = buffer
        self.model = model
        self.table_id = table_id
        self.row_num = row_num
        self.col_num = col_num

        self.d128 = None
        self.double = None
        self.seconds = None
        self.string_id = None
        self.rich_id = None
        self.cell_style_id = None
        self.text_style_id = None
        # self.cond_style_id = None
        # self.cond_rule_style_id = None
        self.formula_id = None
        # self.control_id = None
        self.formula_error_id = None
        self.suggest_id = None
        self.num_format_id = None
        self.currency_format_id = None
        self.date_format_id = None
        self.duration_format_id = None
        self.text_format_id = None
        self.bool_format_id = None
        # self.comment_id = None
        # self.import_warning_id = None
        self.is_currency = False

        if buffer is None:
            return

        version = buffer[0]
        if version != 5:
            raise UnsupportedError(f"Cell storage version {version} is unsupported")

        offset = 12
        flags = unpack("<i", buffer[8:12])[0]

        if flags & 0x1:
            self.d128 = unpack_decimal128(buffer[offset : offset + 16])
            offset += 16
        if flags & 0x2:
            self.double = unpack("<d", buffer[offset : offset + 8])[0]
            offset += 8
        if flags & 0x4:
            self.seconds = unpack("<d", buffer[offset : offset + 8])[0]
            offset += 8
        if flags & 0x8:
            self.string_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x10:
            self.rich_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x20:
            self.cell_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x40:
            self.text_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x80:
            # self.cond_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x100:
        #     self.cond_rule_style_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        if flags & 0x200:
            self.formula_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x400:
        #     self.control_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        # if flags & 0x800:
        #     self.formula_error_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        if flags & 0x1000:
            self.suggest_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # Skip unused flags
        offset += 4 * bin(flags & 0xD00).count("1")
        #
        if flags & 0x2000:
            self.num_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x4000:
            self.currency_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x8000:
            self.date_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x10000:
            self.duration_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x20000:
            self.text_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x40000:
            self.bool_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x80000:
        #     self.comment_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        # if flags & 0x100000:
        #     self.import_warning_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4

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
        elif cell_type == CURRENCY_CELL_TYPE:
            self.value = self.d128
            self.is_currency = True
            self.type = CellType.NUMBER
        else:
            raise UnsupportedError(f"Cell type ID {cell_type} is not recognised")

        if logging.getLogger(__package__).level == logging.DEBUG:
            # Guard to reduce expense of computing fields
            extras = unpack("<H", buffer[6:8])[0]
            table_name = model.table_name(table_id)
            sheet_name = model.sheet_name(model.table_id_to_sheet_id(table_id))
            fields = [
                f"{x}=" + str(getattr(self, x)) if getattr(self, x) is not None else None
                for x in self.__slots__
                if x.endswith("_id")
            ]
            fields = ", ".join([x for x in fields if x if not None])
            debug(
                "%s@%s@[%d,%d]: table_id=%d, type=%s, value=%s, flags=%08x, extras=%04x, %s",
                sheet_name,
                table_name,
                row_num,
                col_num,
                table_id,
                self.type.name,
                self.value,
                flags,
                extras,
                fields,
            )

    def update_value(self, value, cell: object) -> None:
        if cell._type == TSTArchives.numberCellType:
            self.d128 = value
            self.type = CellType.NUMBER
        elif cell._type == TSTArchives.dateCellType:
            self.datetime = value
            self.type = CellType.DATE
        elif cell._type == TSTArchives.durationCellType:
            self.double = value
        self.value = value

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
    @cache(num_args=0)
    def image_data(self) -> Tuple[bytes, str]:
        """Return the background image data for a cell or None if no image."""
        if self.cell_style_id is None:
            return None
        style = self.model.table_style(self.table_id, self.cell_style_id)
        if not style.cell_properties.cell_fill.HasField("image"):
            return None

        image_id = style.cell_properties.cell_fill.image.imagedata.identifier
        datas = self.model.objects[PACKAGE_ID].datas
        stored_filename = [x.file_name for x in datas if x.identifier == image_id][0]
        preferred_filename = [x.preferred_file_name for x in datas if x.identifier == image_id][0]
        all_paths = self.model.objects.file_store.keys()
        image_pathnames = [x for x in all_paths if x == f"Data/{stored_filename}"]
        if len(image_pathnames) == 0:
            warn(
                f"Cannot find file '{preferred_filename}' in Numbers archive",
                RuntimeWarning,
                stacklevel=3,
            )
        else:
            return (self.model.objects.file_store[image_pathnames[0]], preferred_filename)

    def custom_format(self) -> str:  # noqa: PLR0911
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
                formatted_value = format_fraction(self.d128, custom_format)
            elif custom_format.format_type == FormatType.CUSTOM_TEXT:
                formatted_value = decode_text_format(
                    custom_format,
                    self.model.table_string(self.table_id, self.string_id),
                )
            else:
                formatted_value = decode_number_format(
                    custom_format, self.d128, format_map[format_uuid].name
                )
        elif format.format_type == FormatType.DECIMAL:
            return format_decimal(self.d128, format)
        elif format.format_type == FormatType.CURRENCY:
            return format_currency(self.d128, format)
        elif format.format_type == FormatType.BOOLEAN:
            return "TRUE" if self.value else "FALSE"
        elif format.format_type == FormatType.PERCENT:
            return format_decimal(self.d128 * 100, format, percent=True)
        elif format.format_type == FormatType.BASE:
            return format_base(self.d128, format)
        elif format.format_type == FormatType.FRACTION:
            return format_fraction(self.d128, format)
        elif format.format_type == FormatType.SCIENTIFIC:
            return format_scientific(self.d128, format)
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
                formatted_value = decode_date_format(custom_format_string, self.datetime)
            else:
                warn(
                    f"Unexpected custom format type {custom_format.format_type}",
                    UnsupportedWarning,
                    stacklevel=3,
                )
                return ""
        else:
            formatted_value = decode_date_format(format.date_time_format, self.datetime)
        return formatted_value

    def duration_format(self) -> str:
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
                dstr.append(str(dd) + unit_format("millisecond", dd, duration_style, "ms"))
        duration_str = (":" if duration_style == 0 else " ").join(dstr)
        if duration_style == DurationStyle.COMPACT:
            duration_str = re.sub(r":(\d\d\d)$", r".\1", duration_str)

        return duration_str

    def _set_formatting(
        self, format_id: int, format_type: Union[FormattingType, CustomFormattingType]
    ) -> None:
        if format_type == FormattingType.CURRENCY:
            self.currency_format_id = format_id
            self.is_currency = True
        elif format_type in [FormattingType.DATETIME, CustomFormattingType.DATETIME]:
            self.date_format_id = format_id
        else:
            self.num_format_id = format_id


def unpack_decimal128(buffer: bytearray) -> float:
    exp = (((buffer[15] & 0x7F) << 7) | (buffer[14] >> 1)) - 0x1820
    mantissa = buffer[14] & 1
    for i in range(13, -1, -1):
        mantissa = mantissa * 256 + buffer[i]
    sign = 1 if buffer[15] & 0x80 else 0
    if sign == 1:
        mantissa = -mantissa
    value = mantissa * 10**exp
    return float(value)


def decode_date_format_field(field: str, value: datetime) -> str:
    if field in DATETIME_FIELD_MAP:
        s = DATETIME_FIELD_MAP[field]
        if callable(s):
            return s(value)
        else:
            return value.strftime(s)
    else:
        warn(f"Unsupported field code '{field}'", UnsupportedWarning, stacklevel=4)
        return ""


def decode_date_format(format, value):
    """Parse a custom date format string and return a formatted datetime value."""
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


def decode_text_format(format, value: str):
    """Parse a custom date format string and return a formatted number value."""
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


def decode_number_format(format, value, name):  # noqa: PLR0912
    """Parse a custom date format string and return a formatted number value."""
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

    if (match := re.search(r"([#0.,]+(E[+]\d+)?)", custom_format_string)) is None:
        warn(
            f"Can't parse format string '{custom_format_string}'; skipping",
            UnsupportedWarning,
            stacklevel=1,
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
                num_commas = int(math.floor(math.log10(integer)) / 3) if integer != 0 else 0
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

    # value_1 = str(value).split(".")[0]
    # value_2 = sigfig.round(str(value).split(".")[1], sigfig=MAX_SIGNIFICANT_DIGITS, warn=False)
    # int_pad_space_as_zero = (
    #     num_integers > 0
    #     and num_decimals > 0
    #     and int_pad == CellPadding.SPACE
    #     and dec_pad is None
    #     and num_integers > len(value_1)
    #     and num_decimals > len(value_2)
    # )
    int_pad_space_as_zero = False

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
    elif int_pad_space_as_zero or int_pad == CellPadding.ZERO:
        if format.show_thousands_separator:
            formatted_value = f"{integer:0{int_width},}"
        else:
            formatted_value = f"{integer:0{int_width}}"
    elif int_pad == CellPadding.SPACE:
        if format.show_thousands_separator:
            formatted_value = f"{integer:,}".rjust(int_width)
        else:
            formatted_value = str(integer).rjust(int_width)
    elif format.show_thousands_separator:
        formatted_value = f"{integer:,}"
    else:
        formatted_value = str(integer)

    if num_decimals:
        if dec_pad == CellPadding.ZERO or (dec_pad == CellPadding.SPACE and num_integers == 0):
            formatted_value += "." + f"{decimal:,.{dec_width}f}"[2:]
        elif dec_pad == CellPadding.SPACE and decimal == 0 and num_integers > 0:
            formatted_value += ".".ljust(dec_width + 1)
        elif dec_pad == CellPadding.SPACE:
            decimal_str = str(decimal)[2:]
            formatted_value += "." + decimal_str.ljust(dec_width)
        elif decimal or num_integers == 0:
            formatted_value += "." + str(decimal)[2:]

    formatted_value = custom_format_string.replace(format_spec, formatted_value)
    return expand_quotes(formatted_value)


def format_decimal(value: float, format, percent: bool = False) -> str:
    if value is None:
        return ""
    if value < 0 and format.negative_style == 1:
        accounting_style = False
        value = -value
    elif value < 0 and format.negative_style >= 2:
        accounting_style = True
        value = -value
    else:
        accounting_style = False
    thousands = "," if format.show_thousands_separator else ""

    if value.is_integer() and format.decimal_places >= DECIMAL_PLACES_AUTO:
        formatted_value = f"{int(value):{thousands}}"
    else:
        if format.decimal_places >= DECIMAL_PLACES_AUTO:
            formatted_value = str(sigfig.round(value, MAX_SIGNIFICANT_DIGITS, warn=False))
        else:
            formatted_value = sigfig.round(value, MAX_SIGNIFICANT_DIGITS, type=str, warn=False)
            formatted_value = sigfig.round(
                formatted_value, decimals=format.decimal_places, type=str
            )
        if format.show_thousands_separator:
            formatted_value = sigfig.round(formatted_value, spacer=",", spacing=3, type=str)
            try:
                (integer, decimal) = formatted_value.split(".")
                formatted_value = integer + "." + decimal.replace(",", "")
            except ValueError:
                pass

    if percent:
        formatted_value += "%"

    if accounting_style:
        return f"({formatted_value})"
    else:
        return formatted_value


def format_currency(value: float, format) -> str:
    formatted_value = format_decimal(value, format)
    if format.currency_code in CURRENCY_SYMBOLS:
        symbol = CURRENCY_SYMBOLS[format.currency_code]
    else:
        symbol = format.currency_code + " "
    if format.use_accounting_style and value < 0:
        return f"{symbol}\t({formatted_value[1:]})"
    elif format.use_accounting_style:
        return f"{symbol}\t{formatted_value}"
    else:
        return symbol + formatted_value


INT_TO_BASE_CHAR = [str(x) for x in range(0, 10)] + [chr(x) for x in range(ord("A"), ord("Z") + 1)]


def invert_bit_str(value: str) -> str:
    """Invert a binary value"""
    return "".join(["0" if b == "1" else "1" for b in value])


def twos_complement(value: int, base: int) -> str:
    """Calculate the twos complement of a negative integer with minimum 32-bit precision"""
    num_bits = max([32, math.ceil(math.log2(abs(value))) + 1])
    bin_value = bin(abs(value))[2:]
    inverted_bin_value = invert_bit_str(bin_value).rjust(num_bits, "1")
    twos_complement_dec = int(inverted_bin_value, 2) + 1

    if base == 2:
        return bin(twos_complement_dec)[2:].rjust(num_bits, "1")
    elif base == 8:
        return oct(twos_complement_dec)[2:]
    else:
        return hex(twos_complement_dec)[2:].upper()


def format_base(value: float, format) -> str:
    if value == 0:
        return "0".zfill(format.base_places)

    value = round(value)

    is_negative = False
    if not format.base_use_minus_sign and format.base in [2, 8, 16]:
        if value < 0:
            return twos_complement(value, format.base)
        else:
            value = abs(value)
    elif value < 0:
        is_negative = True
        value = abs(value)

    formatted_value = []
    while value:
        formatted_value.append(int(value % format.base))
        value //= format.base
    formatted_value = "".join([INT_TO_BASE_CHAR[x] for x in formatted_value[::-1]])

    if is_negative:
        return "-" + formatted_value.zfill(format.base_places)
    else:
        return formatted_value.zfill(format.base_places)


def format_fraction_parts_to(whole: int, numerator: int, denominator: int):
    if whole > 0:
        if numerator == 0:
            return str(whole)
        else:
            return f"{whole} {numerator}/{denominator}"
    elif numerator == 0:
        return "0"
    elif numerator == denominator:
        return "1"
    return f"{numerator}/{denominator}"


def float_to_fraction(value: float, denominator: int) -> str:
    """Convert a float to the nearest fraction and return as a string."""
    whole = int(value)
    numerator = round(denominator * (value - whole))
    return format_fraction_parts_to(whole, numerator, denominator)


def float_to_n_digit_fraction(value: float, max_digits: int) -> str:
    """Convert a float to a fraction of a maxinum number of digits
    and return as a string.
    """
    max_denominator = 10**max_digits - 1
    (numerator, denominator) = (
        Fraction.from_float(value).limit_denominator(max_denominator).as_integer_ratio()
    )
    whole = int(value)
    numerator -= whole * denominator
    return format_fraction_parts_to(whole, numerator, denominator)


def format_fraction(value: float, format) -> str:
    accuracy = format.fraction_accuracy
    if accuracy & 0xFF000000:
        num_digits = 0x100000000 - accuracy
        return float_to_n_digit_fraction(value, num_digits)
    else:
        return float_to_fraction(value, accuracy)


def format_scientific(value: float, format) -> str:
    formatted_value = sigfig.round(value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
    return f"{formatted_value:.{format.decimal_places}E}"


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
