import logging
import math
import re
from collections import namedtuple
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime as builtin_datetime
from datetime import timedelta as builtin_timedelta
from enum import IntEnum
from fractions import Fraction
from hashlib import sha1
from os.path import basename
from struct import pack, unpack
from typing import Any, List, Optional, Tuple, Union
from warnings import warn

import sigfig
from pendulum import DateTime, Duration, datetime, duration
from pendulum import instance as pendulum_instance

from numbers_parser import __name__ as numbers_parser_name

# from numbers_parser.cell_storage import CellStorage, CellType
from numbers_parser.constants import (
    CHECKBOX_FALSE_VALUE,
    CHECKBOX_TRUE_VALUE,
    CURRENCY_CELL_TYPE,
    CUSTOM_TEXT_PLACEHOLDER,
    DATETIME_FIELD_MAP,
    DECIMAL_PLACES_AUTO,
    DEFAULT_ALIGNMENT,
    DEFAULT_BORDER_COLOR,
    DEFAULT_BORDER_STYLE,
    DEFAULT_BORDER_WIDTH,
    DEFAULT_DATETIME_FORMAT,
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_TEXT_INSET,
    DEFAULT_TEXT_WRAP,
    EMPTY_STORAGE_BUFFER,
    EPOCH,
    MAX_BASE,
    MAX_SIGNIFICANT_DIGITS,
    PACKAGE_ID,
    SECONDS_IN_DAY,
    SECONDS_IN_HOUR,
    SECONDS_IN_WEEK,
    STAR_RATING_VALUE,
    CellPadding,
    CellType,
    ControlFormattingType,
    CustomFormattingType,
    DurationStyle,
    DurationUnits,
    FormattingType,
    FormatType,
    FractionAccuracy,
    NegativeNumberStyle,
    PaddingType,
)
from numbers_parser.currencies import CURRENCIES, CURRENCY_SYMBOLS
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated.TSWPArchives_pb2 import (
    ParagraphStylePropertiesArchive as ParagraphStyle,
)
from numbers_parser.numbers_cache import Cacheable, cache
from numbers_parser.numbers_uuid import NumbersUUID

logger = logging.getLogger(numbers_parser_name)
debug = logger.debug


__all__ = [
    "Alignment",
    "BackgroundImage",
    "BoolCell",
    "Border",
    "BorderType",
    "BulletedTextCell",
    "Cell",
    "CellBorder",
    "CustomFormatting",
    "DateCell",
    "DurationCell",
    "EmptyCell",
    "ErrorCell",
    "Formatting",
    "HorizontalJustification",
    "MergeAnchor",
    "MergeReference",
    "MergedCell",
    "NumberCell",
    "RichTextCell",
    "RGB",
    "Style",
    "TextCell",
    "VerticalJustification",
    "xl_cell_to_rowcol",
    "xl_col_to_name",
    "xl_range",
    "xl_rowcol_to_cell",
]


class BackgroundImage:
    """A named document style that can be applied to cells.

    .. code-block:: python

        fh = open("cats.png", mode="rb")
        image_data = fh.read()
        cats_bg = doc.add_style(
            name="Cats",
            bg_image=BackgroundImage(image_data, "cats.png")
        )
        table.write(0, 0, "❤️ cats", style=cats_bg)

    Currently only standard image files and not 'advanced' image fills are
    supported. Tiling and scaling is not reported back and cannot be changed
    when saving new cells.

    Parameters
    ----------
    data: bytes
        Raw image data for a cell background image.
    filename: str
        Path to the image file.
    """

    def __init__(self, data: Optional[bytes] = None, filename: Optional[str] = None) -> None:
        self._data = data
        self._filename = basename(filename)

    @property
    def data(self) -> bytes:
        """bytes: The background image as bytes for a cell, or None if no image."""
        return self._data

    @property
    def filename(self) -> str:
        """str: The image filename for a cell, or None if no image."""
        return self._filename


class HorizontalJustification(IntEnum):
    LEFT = ParagraphStyle.TextAlignmentType.TATvalue0
    RIGHT = ParagraphStyle.TextAlignmentType.TATvalue1
    CENTER = ParagraphStyle.TextAlignmentType.TATvalue2
    JUSTIFIED = ParagraphStyle.TextAlignmentType.TATvalue3
    AUTO = ParagraphStyle.TextAlignmentType.TATvalue4


class VerticalJustification(IntEnum):
    TOP = ParagraphStyle.DeprecatedParagraphBorderType.PBTvalue0
    MIDDLE = ParagraphStyle.DeprecatedParagraphBorderType.PBTvalue1
    BOTTOM = ParagraphStyle.DeprecatedParagraphBorderType.PBTvalue2


HORIZONTAL_MAP = {
    "left": HorizontalJustification.LEFT,
    "right": HorizontalJustification.RIGHT,
    "center": HorizontalJustification.CENTER,
    "justified": HorizontalJustification.JUSTIFIED,
    "auto": HorizontalJustification.AUTO,
}

VERTICAL_MAP = {
    "top": VerticalJustification.TOP,
    "middle": VerticalJustification.MIDDLE,
    "bottom": VerticalJustification.BOTTOM,
}

_Alignment = namedtuple("Alignment", ["horizontal", "vertical"])


class Alignment(_Alignment):
    def __new__(cls, horizontal=DEFAULT_ALIGNMENT[0], vertical=DEFAULT_ALIGNMENT[1]):
        if isinstance(horizontal, str):
            horizontal = horizontal.lower()
            if horizontal not in HORIZONTAL_MAP:
                msg = "invalid horizontal alignment"
                raise TypeError(msg)
            horizontal = HORIZONTAL_MAP[horizontal]

        if isinstance(vertical, str):
            vertical = vertical.lower()
            if vertical not in VERTICAL_MAP:
                msg = "invalid vertical alignment"
                raise TypeError(msg)
            vertical = VERTICAL_MAP[vertical]

        return super(_Alignment, cls).__new__(cls, (horizontal, vertical))


DEFAULT_ALIGNMENT_CLASS = Alignment(*DEFAULT_ALIGNMENT)

RGB = namedtuple("RGB", ["r", "g", "b"])


@dataclass
class Style:
    """A named document style that can be applied to cells.

    Parameters
    ----------
    alignment: Alignment, optional, default: Alignment("auto", "top")
        Horizontal and vertical alignment of the cell
    bg_color: RGB | List[RGB], optional, default: RGB(0, 0, 0)
        Background color or list of colors for gradients
    bold: bool, optional, default: False
        ``True`` if the cell font is bold
    font_color: RGB, optional, default: RGB(0, 0, 0)) – Font color
    font_size: float, optional, default: DEFAULT_FONT_SIZE
        Font size in points
    font_name: str, optional, default: DEFAULT_FONT_SIZE
        Font name
    italic: bool, optional, default: False
        ``True`` if the cell font is italic
    name: str, optional
        Style name
    underline: bool, optional, default: False) – True if the
        cell font is underline
    strikethrough: bool, optional, default: False) – True if
        the cell font is strikethrough
    first_indent: float, optional, default: 0.0) – First line
        indent in points
    left_indent: float, optional, default: 0.0
        Left indent in points
    right_indent: float, optional, default: 0.0
        Right indent in points
    text_inset: float, optional, default: DEFAULT_TEXT_INSET
        Text inset in points
    text_wrap: str, optional, default: True
        ``True`` if text wrapping is enabled

    Raises
    ------
    TypeError:
        If arguments do not match the specified type or for objects have invalid arguments
    IndexError:
        If an image filename already exists in document
    """

    alignment: Alignment = DEFAULT_ALIGNMENT_CLASS  # : horizontal and vertical alignment
    bg_image: object = None  # : backgroung image
    bg_color: Union[RGB, List[RGB]] = None
    font_color: RGB = RGB(0, 0, 0)
    font_size: float = DEFAULT_FONT_SIZE
    font_name: str = DEFAULT_FONT
    bold: bool = False
    italic: bool = False
    strikethrough: bool = False
    underline: bool = False
    first_indent: float = 0
    left_indent: float = 0
    right_indent: float = 0
    text_inset: float = DEFAULT_TEXT_INSET
    text_wrap: bool = DEFAULT_TEXT_WRAP
    name: str = None
    _text_style_obj_id: int = None
    _cell_style_obj_id: int = None
    _update_cell_style: bool = False
    _update_text_style: bool = False

    @staticmethod
    def _text_attrs():
        return [
            "alignment",
            "bold",
            "first_indent",
            "font_color",
            "font_name",
            "font_size",
            "italic",
            "left_indent",
            "name",
            "right_indent",
            "strikethrough",
            "text_inset",
            "underline",
        ]

    @staticmethod
    def _cell_attrs():
        return [
            "alignment",
            "bg_color",
            "bg_image",
            "first_indent",
            "left_indent",
            "right_indent",
            "text_inset",
            "text_wrap",
        ]

    @classmethod
    def from_storage(cls, cell: object, model: object):
        bg_image = BackgroundImage(*cell._image_data) if cell._image_data is not None else None
        return Style(
            alignment=model.cell_alignment(cell),
            bg_image=bg_image,
            bg_color=model.cell_bg_color(cell),
            font_color=model.cell_font_color(cell),
            font_size=model.cell_font_size(cell),
            font_name=model.cell_font_name(cell),
            bold=model.cell_is_bold(cell),
            italic=model.cell_is_italic(cell),
            strikethrough=model.cell_is_strikethrough(cell),
            underline=model.cell_is_underline(cell),
            name=model.cell_style_name(cell),
            first_indent=model.cell_first_indent(cell),
            left_indent=model.cell_left_indent(cell),
            right_indent=model.cell_right_indent(cell),
            text_inset=model.cell_text_inset(cell),
            text_wrap=model.cell_text_wrap(cell),
            _text_style_obj_id=model.text_style_object_id(cell),
            _cell_style_obj_id=model.cell_style_object_id(cell),
        )

    def __post_init__(self):
        self.bg_color = rgb_color(self.bg_color)
        self.font_color = rgb_color(self.font_color)

        if not isinstance(self.font_size, float):
            msg = "size must be a float number of points"
            raise TypeError(msg)
        if not isinstance(self.font_name, str):
            msg = "font name must be a string"
            raise TypeError(msg)

        for attr in ["bold", "italic", "underline", "strikethrough"]:
            if not isinstance(getattr(self, attr), bool):
                msg = f"{attr} argument must be boolean"
                raise TypeError(msg)

    def __setattr__(self, name: str, value: Any) -> None:
        """Detect changes to cell styles and flag the style for
        possible updates when saving the document.
        """
        if name in ["bg_color", "font_color"]:
            value = rgb_color(value)
        if name == "alignment":
            value = alignment(value)
        if name in Style._text_attrs():
            self.__dict__["_update_text_style"] = True
        if name in Style._cell_attrs():
            self.__dict__["_update_cell_style"] = True

        if name not in ["_update_text_style", "_update_cell_style"]:
            self.__dict__[name] = value


def rgb_color(color) -> RGB:
    """Raise a TypeError if a color is not a valid RGB value."""
    if color is None:
        return None
    if isinstance(color, RGB):
        return color
    if isinstance(color, tuple):
        if not (len(color) == 3 and all(isinstance(x, int) for x in color)):
            msg = "RGB color must be an RGB or a tuple of 3 integers"
            raise TypeError(msg)
        return RGB(*color)
    elif isinstance(color, list):
        return [rgb_color(c) for c in color]
    msg = "RGB color must be an RGB or a tuple of 3 integers"
    raise TypeError(msg)


def alignment(value) -> Alignment:
    """Raise a TypeError if a alignment is not a valid."""
    if value is None:
        return Alignment()
    if isinstance(value, Alignment):
        return value
    if isinstance(value, tuple):
        if not (len(value) == 2 and all(isinstance(x, (int, str)) for x in value)):
            msg = "Alignment must be an Alignment or a tuple of 2 integers/strings"
            raise TypeError(msg)
        return Alignment(*value)
    msg = "Alignment must be an Alignment or a tuple of 2 integers/strings"
    raise TypeError(msg)


BORDER_STYLE_MAP = {"solid": 0, "dashes": 1, "dots": 2, "none": 3}


class BorderType(IntEnum):
    SOLID = BORDER_STYLE_MAP["solid"]
    DASHES = BORDER_STYLE_MAP["dashes"]
    DOTS = BORDER_STYLE_MAP["dots"]
    NONE = BORDER_STYLE_MAP["none"]


class Border:
    """Create a cell border to use with the :py:class:`~numbers_parser.Table` method
    :py:meth:`~numbers_parser.Table.set_cell_border`.

    .. code-block:: python

        border_style = Border(8.0, RGB(29, 177, 0)
        table.set_cell_border("B6", "left", border_style, "solid"), 3)
        table.set_cell_border(6, 1, "right", border_style, "dashes"))

    Parameters
    ----------
    width: float, optional, default: 0.35
        Number of rows in the first table of a new document.
    color: RGB, optional, default: RGB(0, 0, 0)
        The line color for the border if present
    style: BorderType, optional, default: ``None``
        The type of border to create or ``None`` if there is no border defined. Valid
        border types are:

        * ``"solid"``: a solid line
        * ``"dashes"``: a dashed line
        * ``"dots"``: a dotted line

    Raises
    ------
    TypeError:
        If the width is not a float, or the border type is invalid.
    """

    def __init__(
        self,
        width: float = DEFAULT_BORDER_WIDTH,
        color: RGB = None,
        style: BorderType = None,
        _order: int = 0,
    ) -> None:
        if not isinstance(width, float):
            msg = "width must be a float number of points"
            raise TypeError(msg)
        self.width = width

        if color is None:
            color = RGB(*DEFAULT_BORDER_COLOR)
        self.color = rgb_color(color)

        if style is None:
            style = BorderType(BORDER_STYLE_MAP[DEFAULT_BORDER_STYLE])
        if isinstance(style, str):
            style = style.lower()
            if style not in BORDER_STYLE_MAP:
                msg = "invalid border style"
                raise TypeError(msg)
            self.style = BORDER_STYLE_MAP[style]
        else:
            self.style = style

        self._order = _order

    def __str__(self) -> str:
        style_name = BorderType(self.style).name.lower()
        return f"Border(width={self.width}, color={self.color}, style={style_name})"

    def __eq__(self, value: object) -> bool:
        return all(
            [self.width == value.width, self.color == value.color, self.style == value.style],
        )


class CellBorder:
    def __init__(
        self,
        top_merged: bool = False,
        right_merged: bool = False,
        bottom_merged: bool = False,
        left_merged: bool = False,
    ) -> None:
        self._top = None
        self._right = None
        self._bottom = None
        self._left = None
        self._top_merged = top_merged
        self._right_merged = right_merged
        self._bottom_merged = bottom_merged
        self._left_merged = left_merged

    @property
    def top(self):
        if self._top_merged:
            return None
        elif self._top is None:
            return None
        return self._top

    @top.setter
    def top(self, value):
        if self._top is None:
            self._top = value
        elif value._order > self.top._order:
            self._top = value

    @property
    def right(self):
        if self._right_merged:
            return None
        elif self._right is None:
            return None
        return self._right

    @right.setter
    def right(self, value):
        if self._right is None:
            self._right = value
        elif value._order > self._right._order:
            self._right = value

    @property
    def bottom(self):
        if self._bottom_merged:
            return None
        elif self._bottom is None:
            return None
        return self._bottom

    @bottom.setter
    def bottom(self, value):
        if self._bottom is None:
            self._bottom = value
        elif value._order > self._bottom._order:
            self._bottom = value

    @property
    def left(self):
        if self._left_merged:
            return None
        elif self._left is None:
            return None
        return self._left

    @left.setter
    def left(self, value):
        if self._left is None:
            self._left = value
        elif value._order > self._left._order:
            self._left = value


class MergeReference:
    """Cell reference for cells eliminated by a merge."""

    def __init__(self, row_start: int, col_start: int, row_end: int, col_end: int) -> None:
        self.rect = (row_start, col_start, row_end, col_end)


class MergeAnchor:
    """Cell reference for the merged cell."""

    def __init__(self, size: Tuple) -> None:
        self.size = size


@dataclass
class CellStorageFlags:
    _string_id: int = None
    _rich_id: int = None
    _cell_style_id: int = None
    _text_style_id: int = None
    # _cond_style_id: int = None
    # _cond_rule_style_id: int = None
    _formula_id: int = None
    _control_id: int = None
    _formula_error_id: int = None
    _suggest_id: int = None
    _num_format_id: int = None
    _currency_format_id: int = None
    _date_format_id: int = None
    _duration_format_id: int = None
    _text_format_id: int = None
    _bool_format_id: int = None
    # _comment_id: int = None
    # _import_warning_id: int = None

    def __str__(self) -> str:
        fields = [
            f"{k[1:]}={v}" for k, v in asdict(self).items() if k.endswith("_id") and v is not None
        ]
        return ", ".join([x for x in fields if x if not None])

    def flags(self):
        return [x.name for x in fields(self)]


class Cell(CellStorageFlags, Cacheable):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value) -> None:
        self._value = value
        self.row = row
        self.col = col
        self._is_bulleted = False
        self._storage = None
        self._style = None
        self._d128 = None
        self._double = None
        self._seconds = None
        super().__init__()

    def __str__(self) -> str:
        table_name = self._model.table_name(self._table_id)
        sheet_name = self._model.sheet_name(self._model.table_id_to_sheet_id(self._table_id))
        cell_str = f"{sheet_name}@{table_name}[{self.row},{self.col}]:"
        cell_str += f"table_id={self._table_id}, type={self._type.name}, "
        cell_str += f"value={self._value}, flags={self._flags:08x}, extras={self._extras:04x}"
        return ", ".join([cell_str, super().__str__()])

    @property
    def image_filename(self):
        warn(
            "image_filename is deprecated and will be removed in the future. "
            + "Please use the style property",
            DeprecationWarning,
            stacklevel=2,
        )
        if self.style is not None and self.style.bg_image is not None:
            return self.style.bg_image.filename
        else:
            return None

    @property
    def image_data(self):
        warn(
            "image_data is deprecated and will be removed in the future. "
            + "Please use the style property",
            DeprecationWarning,
            stacklevel=2,
        )
        if self.style is not None and self.style.bg_image is not None:
            return self.style.bg_image.data
        else:
            return None

    @property
    def is_formula(self) -> bool:
        """bool: ``True`` if the cell contains a formula."""
        table_formulas = self._model.table_formulas(self._table_id)
        return table_formulas.is_formula(self.row, self.col)

    @property
    @cache(num_args=0)
    def formula(self) -> str:
        """str: The formula in a cell.

        Formula evaluation relies on Numbers storing current values which should
        usually be the case. In cells containing a formula, :py:meth:`numbers_parser.Cell.value`
        returns computed value of the formula.

        Returns
        -------
            str:
                The text of the foruma in a cell, or `None` if there is no formula
                present in a cell.
        """
        if self._formula_id is not None:
            table_formulas = self._model.table_formulas(self._table_id)
            return table_formulas.formula(self._formula_id, self.row, self.col)
        else:
            return None

    @property
    def is_bulleted(self) -> bool:
        """bool: ``True`` if the cell contains text bullets."""
        return self._is_bulleted

    @property
    def bullets(self) -> Union[List[str], None]:
        r"""List[str] | None: The bullets in a cell, or ``None``.

        Cells that contain bulleted or numbered lists are identified
        by :py:attr:`numbers_parser.Cell.is_bulleted`. For these cells,
        :py:attr:`numbers_parser.Cell.value` returns the whole cell contents.
        Bullets can also be extracted into a list of paragraphs cell without the
        bullet or numbering character. Newlines are not included in the
        bullet list.

        Example:
        -------
        .. code-block:: python

            doc = Document("bullets.numbers")
            sheets = doc.sheets
            tables = sheets[0].tables
            table = tables[0]
            if not table.cell(0, 1).is_bulleted:
                print(table.cell(0, 1).value)
            else:
                bullets = ["* " + s for s in table.cell(0, 1).bullets]
                print("\n".join(bullets))
                    return None
        """
        return None

    @property
    def formatted_value(self) -> str:
        """str: The formatted value of the cell as it appears in Numbers.

        Interactive elements are converted into a suitable text format where
        supported, or as their number values where there is no suitable
        visual representation. Currently supported mappings are:

        * Checkboxes are U+2610 (Ballow Box) or U+2611 (Ballot Box with Check)
        * Ratings are their star value represented using (U+2605) (Black Star)

        .. code-block:: python

            >>> table = doc.default_table
            >>> table.cell(0,0).value
            False
            >>> table.cell(0,0).formatted_value
            '☐'
            >>> table.cell(0,1).value
            True
            >>> table.cell(0,1).formatted_value
            '☑'
            >>> table.cell(1,1).value
            3.0
            >>> table.cell(1,1).formatted_value
            '★★★'
        """
        if self._duration_format_id is not None and self._double is not None:
            return self._duration_format()
        elif self._date_format_id is not None and self._seconds is not None:
            return self._date_format()
        elif (
            self._text_format_id is not None
            or self._num_format_id is not None
            or self._currency_format_id is not None
            or self._bool_format_id is not None
        ):
            return self._custom_format()
        else:
            return str(self.value)

    @property
    def style(self) -> Union[Style, None]:
        """Style | None: The :class:`Style` associated with the cell or ``None``.

        Warns:
        -----
            UnsupportedWarning: On assignment; use
                :py:meth:`numbers_parser.Table.set_cell_style` instead.
        """
        if self._style is None:
            self._style = Style.from_storage(self, self._model)
        return self._style

    @style.setter
    def style(self, _):
        warn(
            "cell style cannot be set; use Table.set_cell_style() instead",
            UnsupportedWarning,
            stacklevel=2,
        )

    @property
    def border(self) -> Union[CellBorder, None]:
        """CellBorder| None: The :class:`CellBorder` associated with the cell or ``None``.

        Warns:
        -----
            UnsupportedWarning: On assignment; use
                :py:meth:`numbers_parser.Table.set_cell_border` instead.
        """
        self._model.extract_strokes(self._table_id)
        return self._border

    @border.setter
    def border(self, _):
        warn(
            "cell border values cannot be set; use Table.set_cell_border() instead",
            UnsupportedWarning,
            stacklevel=2,
        )

    @classmethod
    def _empty_cell(cls, table_id: int, row: int, col: int, model: object):
        return Cell._from_storage(table_id, row, col, EMPTY_STORAGE_BUFFER, model)

    @classmethod
    def _merged_cell(cls, table_id: int, row: int, col: int, model: object):
        cell = MergedCell(row, col)
        cell._model = model
        cell._table_id = table_id
        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row, col)))
        return cell

    @classmethod
    def _from_value(cls, row: int, col: int, value):
        # TODO: write needs to retain/init the border
        if isinstance(value, str):
            cell = TextCell(row, col, value)
        elif isinstance(value, bool):
            cell = BoolCell(row, col, value)
        elif isinstance(value, int):
            cell = NumberCell(row, col, value)
        elif isinstance(value, float):
            rounded_value = sigfig.round(value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
            if rounded_value != value:
                warn(
                    f"'{value}' rounded to {MAX_SIGNIFICANT_DIGITS} significant digits",
                    RuntimeWarning,
                    stacklevel=2,
                )
            cell = NumberCell(row, col, rounded_value)
        elif isinstance(value, (DateTime, builtin_datetime)):
            cell = DateCell(row, col, pendulum_instance(value))
        elif isinstance(value, (Duration, builtin_timedelta)):
            cell = DurationCell(row, col, value)
        else:
            raise ValueError("Can't determine cell type from type " + type(value).__name__)

        return cell

    @classmethod
    def _from_storage(  # noqa: PLR0913, PLR0912
        cls,
        table_id: int,
        row: int,
        col: int,
        buffer: bytearray,
        model: object,
    ) -> None:
        d128 = None
        double = None
        seconds = None

        version = buffer[0]
        if version != 5:
            msg = f"Cell storage version {version} is unsupported"
            raise UnsupportedError(msg)

        offset = 12
        storage_flags = CellStorageFlags()
        flags = unpack("<i", buffer[8:12])[0]

        if flags & 0x1:
            d128 = _unpack_decimal128(buffer[offset : offset + 16])
            offset += 16
        if flags & 0x2:
            double = unpack("<d", buffer[offset : offset + 8])[0]
            offset += 8
        if flags & 0x4:
            seconds = unpack("<d", buffer[offset : offset + 8])[0]
            offset += 8
        if flags & 0x8:
            storage_flags._string_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x10:
            storage_flags._rich_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x20:
            storage_flags._cell_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x40:
            storage_flags._text_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x80:
            # storage_flags._cond_style_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x100:
        #     storage_flags._cond_rule_style_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        if flags & 0x200:
            storage_flags._formula_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x400:
            storage_flags._control_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x800:
        #     storage_flags._formula_error_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        if flags & 0x1000:
            storage_flags._suggest_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # Skip unused flags
        offset += 4 * bin(flags & 0x900).count("1")
        #
        if flags & 0x2000:
            storage_flags._num_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x4000:
            storage_flags._currency_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x8000:
            storage_flags._date_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x10000:
            storage_flags._duration_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x20000:
            storage_flags._text_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        if flags & 0x40000:
            storage_flags._bool_format_id = unpack("<i", buffer[offset : offset + 4])[0]
            offset += 4
        # if flags & 0x80000:
        #     cstorage_flags._omment_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4
        # if flags & 0x100000:
        #     storage_flags._import_warning_id = unpack("<i", buffer[offset : offset + 4])[0]
        #     offset += 4

        cell_type = buffer[1]
        if cell_type == TSTArchives.genericCellType:
            cell = EmptyCell(row, col)
        elif cell_type == TSTArchives.numberCellType:
            cell = NumberCell(row, col, d128)
        elif cell_type == TSTArchives.textCellType:
            cell = TextCell(row, col, model.table_string(table_id, storage_flags._string_id))
        elif cell_type == TSTArchives.dateCellType:
            cell = DateCell(row, col, EPOCH + duration(seconds=seconds))
            cell._datetime = cell._value
        elif cell_type == TSTArchives.boolCellType:
            cell = BoolCell(row, col, double > 0.0)
        elif cell_type == TSTArchives.durationCellType:
            cell = DurationCell(row, col, duration(seconds=double))
        elif cell_type == TSTArchives.formulaErrorCellType:
            cell = ErrorCell(row, col)
        elif cell_type == TSTArchives.automaticCellType:
            cell = RichTextCell(row, col, model.table_rich_text(table_id, storage_flags._rich_id))
        elif cell_type == CURRENCY_CELL_TYPE:
            cell = NumberCell(row, col, d128, cell_type=CellType.CURRENCY)
        else:
            msg = f"Cell type ID {cell_type} is not recognised"
            raise UnsupportedError(msg)

        cell._copy_flags(storage_flags)
        cell._buffer = buffer
        cell._model = model
        cell._table_id = table_id
        cell._d128 = d128
        cell._double = double
        cell._seconds = seconds
        cell._extras = unpack("<H", buffer[6:8])[0]
        cell._flags = flags

        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row, col)))

        if logging.getLogger(__package__).level == logging.DEBUG:
            # Guard to reduce expense of computing fields
            debug(str(cell))

        return cell

    def _copy_flags(self, storage_flags: CellStorageFlags):
        for flag in storage_flags.flags():
            setattr(self, flag, getattr(storage_flags, flag))

    def _set_merge(self, merge_ref):
        if isinstance(merge_ref, MergeAnchor):
            self.is_merged = True
            self.size = merge_ref.size
            self.merge_range = None
            self.rect = None
            self._border = CellBorder()
        elif isinstance(merge_ref, MergeReference):
            self.is_merged = False
            self.size = None
            self.row_start = merge_ref.rect[0]
            self.col_start = merge_ref.rect[1]
            self.row_end = merge_ref.rect[2]
            self.col_end = merge_ref.rect[3]
            self.merge_range = xl_range(*merge_ref.rect)
            self.rect = merge_ref.rect
            top_merged = self.row > self.row_start
            right_merged = self.col < self.col_end
            bottom_merged = self.row < self.row_end
            left_merged = self.col > self.col_start
            self._border = CellBorder(top_merged, right_merged, bottom_merged, left_merged)
        else:
            self.is_merged = False
            self.size = (1, 1)
            self.merge_range = None
            self.rect = None
            self._border = CellBorder()

    def _to_buffer(self) -> bytearray:  # noqa: PLR0912, PLR0915
        """Create a storage buffer for a cell using v5 (modern) layout."""
        if self._style is not None:
            if self._style._text_style_obj_id is not None:
                self._text_style_id = self._model._table_styles.lookup_key(
                    self._table_id,
                    TSPMessages.Reference(identifier=self._style._text_style_obj_id),
                )
                self._model.add_component_reference(
                    self._style._text_style_obj_id,
                    parent_id=self._model._table_styles.id(self._table_id),
                )

            if self._style._cell_style_obj_id is not None:
                self._cell_style_id = self._model._table_styles.lookup_key(
                    self._table_id,
                    TSPMessages.Reference(identifier=self._style._cell_style_obj_id),
                )
                self._model.add_component_reference(
                    self._style._cell_style_obj_id,
                    parent_id=self._model._table_styles.id(self._table_id),
                )

        length = 12
        if isinstance(self, NumberCell):
            flags = 1
            length += 16
            if self._type == CellType.CURRENCY:
                cell_type = CURRENCY_CELL_TYPE
            else:
                cell_type = TSTArchives.numberCellType
            value = _pack_decimal128(self.value)
        elif isinstance(self, TextCell):
            flags = 8
            length += 4
            cell_type = TSTArchives.textCellType
            value = pack("<i", self._model.table_string_key(self._table_id, self.value))
        elif isinstance(self, DateCell):
            flags = 4
            length += 8
            cell_type = TSTArchives.dateCellType
            date_delta = self._value.astimezone() - EPOCH
            value = pack("<d", float(date_delta.total_seconds()))
        elif isinstance(self, BoolCell):
            flags = 2
            length += 8
            cell_type = TSTArchives.boolCellType
            value = pack("<d", float(self.value))
        elif isinstance(self, DurationCell):
            flags = 2
            length += 8
            cell_type = TSTArchives.durationCellType
            value = pack("<d", float(self.value.total_seconds()))
        elif isinstance(self, EmptyCell):
            flags = 0
            cell_type = TSTArchives.emptyCellValueType
            value = b""
        elif isinstance(self, MergedCell):
            return None
        elif isinstance(self, RichTextCell):
            flags = 0
            length += 4
            cell_type = TSTArchives.automaticCellType
            value = pack("<i", self._rich_id)
        else:
            data_type = type(self).__name__
            table_name = self._model.table_name(self._table_id)
            table_ref = f"@{table_name}:[{self.row},{self.col}]"
            warn(
                f"{table_ref}: unsupported data type {data_type} for save",
                UnsupportedWarning,
                stacklevel=1,
            )
            return None

        storage = bytearray(12)
        storage[0] = 5
        storage[1] = cell_type
        storage += value

        if self._rich_id is not None:
            flags |= 0x10
            length += 4
            storage += pack("<i", self._rich_id)
        if self._cell_style_id is not None:
            flags |= 0x20
            length += 4
            storage += pack("<i", self._cell_style_id)
        if self._text_style_id is not None:
            flags |= 0x40
            length += 4
            storage += pack("<i", self._text_style_id)
        if self._formula_id is not None:
            flags |= 0x200
            length += 4
            storage += pack("<i", self._formula_id)
        if self._control_id is not None:
            flags |= 0x400
            length += 4
            storage += pack("<i", self._control_id)
        if self._suggest_id is not None:
            flags |= 0x1000
            length += 4
            storage += pack("<i", self._suggest_id)
        if self._num_format_id is not None:
            flags |= 0x2000
            length += 4
            storage += pack("<i", self._num_format_id)
            storage[6] |= 1
            # storage[6:8] = pack("<h", 1)
        if self._currency_format_id is not None:
            flags |= 0x4000
            length += 4
            storage += pack("<i", self._currency_format_id)
            storage[6] |= 2
        if self._date_format_id is not None:
            flags |= 0x8000
            length += 4
            storage += pack("<i", self._date_format_id)
            storage[6] |= 8
        if self._duration_format_id is not None:
            flags |= 0x10000
            length += 4
            storage += pack("<i", self._duration_format_id)
            storage[6] |= 4
        if self._text_format_id is not None:
            flags |= 0x20000
            length += 4
            storage += pack("<i", self._text_format_id)
        if self._bool_format_id is not None:
            flags |= 0x40000
            length += 4
            storage += pack("<i", self._bool_format_id)
            storage[6] |= 0x20
        if self._string_id is not None:
            storage[6] |= 0x80

        storage[8:12] = pack("<i", flags)
        if len(storage) < 32:
            storage += bytearray(32 - length)

        return storage[0:length]

    def _update_value(self, value, cell: object) -> None:
        if cell._type == CellType.NUMBER:
            self._d128 = value
        elif cell._type == CellType.DATE:
            self._datetime = value
        self._value = value

    @property
    @cache(num_args=0)
    def _image_data(self) -> Tuple[bytes, str]:
        """Return the background image data for a cell or None if no image."""
        if self._cell_style_id is None:
            return None
        style = self._model.table_style(self._table_id, self._cell_style_id)
        if not style.cell_properties.cell_fill.HasField("image"):
            return None

        image_id = style.cell_properties.cell_fill.image.imagedata.identifier
        datas = self._model.objects[PACKAGE_ID].datas
        stored_filename = next(
            x.file_name for x in datas if x.identifier == image_id
        )  # pragma: nocover (issue-1333)
        preferred_filename = next(
            x.preferred_file_name for x in datas if x.identifier == image_id
        )  # pragma: nocover (issue-1333)
        all_paths = self._model.objects.file_store.keys()
        image_pathnames = [x for x in all_paths if x == f"Data/{stored_filename}"]

        if len(image_pathnames) == 0:
            warn(
                f"Cannot find file '{preferred_filename}' in Numbers archive",
                RuntimeWarning,
                stacklevel=3,
            )
            return None
        else:
            image_data = self._model.objects.file_store[image_pathnames[0]]
            digest = sha1(image_data).digest()
            if digest not in self._model._images:
                self._model._images[digest] = image_id

            return (image_data, preferred_filename)

    def _custom_format(self) -> str:  # noqa: PLR0911
        if self._text_format_id is not None and self._type == CellType.TEXT:
            format = self._model.table_format(self._table_id, self._text_format_id)
        elif self._currency_format_id is not None:
            format = self._model.table_format(self._table_id, self._currency_format_id)
        elif self._bool_format_id is not None and self._type == CellType.BOOL:
            format = self._model.table_format(self._table_id, self._bool_format_id)
        elif self._num_format_id is not None:
            format = self._model.table_format(self._table_id, self._num_format_id)
        else:
            return str(self.value)

        debug("custom_format: @[%d,%d]: format_type=%s, ", self.row, self.col, format.format_type)

        if format.HasField("custom_uid"):
            format_uuid = NumbersUUID(format.custom_uid).hex
            format_map = self._model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            if custom_format.requires_fraction_replacement:
                formatted_value = _format_fraction(self._d128, custom_format)
            elif custom_format.format_type == FormatType.CUSTOM_TEXT:
                formatted_value = _decode_text_format(
                    custom_format,
                    self._model.table_string(self._table_id, self._string_id),
                )
            else:
                formatted_value = _decode_number_format(
                    custom_format,
                    self._d128,
                    format_map[format_uuid].name,
                )
        elif format.format_type == FormatType.DECIMAL:
            return _format_decimal(self._d128, format)
        elif format.format_type == FormatType.CURRENCY:
            return _format_currency(self._d128, format)
        elif format.format_type == FormatType.BOOLEAN:
            return "TRUE" if self.value else "FALSE"
        elif format.format_type == FormatType.PERCENT:
            return _format_decimal(self._d128 * 100, format, percent=True)
        elif format.format_type == FormatType.BASE:
            return _format_base(self._d128, format)
        elif format.format_type == FormatType.FRACTION:
            return _format_fraction(self._d128, format)
        elif format.format_type == FormatType.SCIENTIFIC:
            return _format_scientific(self._d128, format)
        elif format.format_type == FormatType.CHECKBOX:
            return CHECKBOX_TRUE_VALUE if self.value else CHECKBOX_FALSE_VALUE
        elif format.format_type == FormatType.RATING:
            return STAR_RATING_VALUE * int(self._d128)
        else:
            formatted_value = str(self.value)
        return formatted_value

    def _date_format(self) -> str:
        format = self._model.table_format(self._table_id, self._date_format_id)
        if format.HasField("custom_uid"):
            format_uuid = NumbersUUID(format.custom_uid).hex
            format_map = self._model.custom_format_map()
            custom_format = format_map[format_uuid].default_format
            custom_format_string = custom_format.custom_format_string
            if custom_format.format_type == FormatType.CUSTOM_DATE:
                formatted_value = _decode_date_format(custom_format_string, self._datetime)
            else:
                warn(
                    f"Unexpected custom format type {custom_format.format_type}",
                    UnsupportedWarning,
                    stacklevel=3,
                )
                return ""
        else:
            formatted_value = _decode_date_format(format.date_time_format, self._datetime)
        return formatted_value

    def _duration_format(self) -> str:
        format = self._model.table_format(self._table_id, self._duration_format_id)
        debug(
            "duration_format: @[%d,%d]: table_id=%d, duration_format_id=%d, duration_style=%s",
            self.row,
            self.col,
            self._table_id,
            self._duration_format_id,
            format.duration_style,
        )

        duration_style = format.duration_style
        unit_largest = format.duration_unit_largest
        unit_smallest = format.duration_unit_smallest
        if format.use_automatic_duration_units:
            unit_smallest, unit_largest = _auto_units(self._double, format)

        d = self._double
        dd = int(self._double)
        dstr = []

        def unit_in_range(largest, smallest, unit_type):
            return largest <= unit_type and smallest >= unit_type

        def pad_digits(d, largest, smallest, unit_type):
            return (largest == unit_type and smallest == unit_type) or d >= 10

        if unit_largest == DurationUnits.WEEK:
            dd = int(d / SECONDS_IN_WEEK)
            if unit_smallest != DurationUnits.WEEK:
                d -= SECONDS_IN_WEEK * dd
            dstr.append(str(dd) + _unit_format("week", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.DAY):
            dd = int(d / SECONDS_IN_DAY)
            if unit_smallest > DurationUnits.DAY:
                d -= SECONDS_IN_DAY * dd
            dstr.append(str(dd) + _unit_format("day", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.HOUR):
            dd = int(d / SECONDS_IN_HOUR)
            if unit_smallest > DurationUnits.HOUR:
                d -= SECONDS_IN_HOUR * dd
            dstr.append(str(dd) + _unit_format("hour", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.MINUTE):
            dd = int(d / 60)
            if unit_smallest > DurationUnits.MINUTE:
                d -= 60 * dd
            if duration_style == DurationStyle.COMPACT:
                pad = pad_digits(dd, unit_smallest, unit_largest, DurationUnits.MINUTE)
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + _unit_format("minute", dd, duration_style))

        if unit_in_range(unit_largest, unit_smallest, DurationUnits.SECOND):
            dd = int(d)
            if unit_smallest > DurationUnits.SECOND:
                d -= dd
            if duration_style == DurationStyle.COMPACT:
                pad = pad_digits(dd, unit_smallest, unit_largest, DurationUnits.SECOND)
                dstr.append(("" if pad else "0") + str(dd))
            else:
                dstr.append(str(dd) + _unit_format("second", dd, duration_style))

        if unit_smallest >= DurationUnits.MILLISECOND:
            dd = int(round(1000 * d))
            if duration_style == DurationStyle.COMPACT:
                padding = "0" if dd >= 10 else "00"
                padding = "" if dd >= 100 else padding
                dstr.append(f"{padding}{dd}")
            else:
                dstr.append(str(dd) + _unit_format("millisecond", dd, duration_style, "ms"))
        duration_str = (":" if duration_style == 0 else " ").join(dstr)
        if duration_style == DurationStyle.COMPACT:
            duration_str = re.sub(r":(\d\d\d)$", r".\1", duration_str)

        return duration_str

    def _set_formatting(
        self,
        format_id: int,
        format_type: Union[FormattingType, CustomFormattingType],
        control_id: Optional[int] = None,
        is_currency: bool = False,
    ) -> None:
        self._is_currency = is_currency
        if format_type == FormattingType.CURRENCY:
            self._currency_format_id = format_id
        elif format_type == FormattingType.TICKBOX:
            self._bool_format_id = format_id
            self._control_id = control_id
        elif format_type == FormattingType.RATING:
            self._num_format_id = format_id
            self._control_id = control_id
        elif format_type in [FormattingType.SLIDER, FormattingType.STEPPER]:
            if is_currency:
                self._currency_format_id = format_id
            else:
                self._num_format_id = format_id
            self._control_id = control_id
        elif format_type == FormattingType.POPUP:
            self._text_format_id = format_id
            self._control_id = control_id
        elif format_type in [FormattingType.DATETIME, CustomFormattingType.DATETIME]:
            self._date_format_id = format_id
        elif format_type in [FormattingType.TEXT, CustomFormattingType.TEXT]:
            self._text_format_id = format_id
        else:
            self._num_format_id = format_id


class NumberCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: float, cell_type=CellType.NUMBER) -> None:
        self._type = cell_type
        super().__init__(row, col, value)

    @property
    def value(self) -> int:
        return self._value


class TextCell(Cell):
    def __init__(self, row: int, col: int, value: str) -> None:
        self._type = CellType.TEXT
        super().__init__(row, col, value)

    @property
    def value(self) -> str:
        return self._value


class RichTextCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value) -> None:
        super().__init__(row, col, value["text"])
        self._type = CellType.RICH_TEXT
        self._bullets = value["bullets"]
        self._hyperlinks = value["hyperlinks"]
        if value["bulleted"]:
            self._formatted_bullets = [
                (
                    value["bullet_chars"][i] + " " + value["bullets"][i]
                    if value["bullet_chars"][i] is not None
                    else value["bullets"][i]
                )
                for i in range(len(self._bullets))
            ]
            self._is_bulleted = True

    @property
    def value(self) -> str:
        return self._value

    @property
    def bullets(self) -> List[str]:
        """List[str]: A list of the text bullets in the cell."""
        return self._bullets

    @property
    def formatted_bullets(self) -> str:
        """str: The bullets as a formatted multi-line string."""
        return self._formatted_bullets

    @property
    def hyperlinks(self) -> Union[List[Tuple], None]:
        """List[Tuple] | None: the hyperlinks in a cell or ``None``.

        Numbers does not support hyperlinks to cells within a spreadsheet, but does
        allow embedding links in cells. When cells contain hyperlinks,
        `numbers_parser` returns the text version of the cell. The `hyperlinks` property
        of cells where :py:attr:`numbers_parser.Cell.is_bulleted` is ``True`` is a
        list of text and URL tuples.

        Example:
        -------
        .. code-block:: python

            cell = table.cell(0, 0)
            (text, url) = cell.hyperlinks[0]
        """
        return self._hyperlinks


# Backwards compatibility to earlier class names
class BulletedTextCell(RichTextCell):
    pass


class EmptyCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int) -> None:
        super().__init__(row, col, None)
        self._type = CellType.EMPTY

    @property
    def value(self):
        return None

    @property
    def formatted_value(self):
        return ""


class BoolCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: bool) -> None:
        super().__init__(row, col, value)
        self._type = CellType.BOOL
        self._value = value

    @property
    def value(self) -> bool:
        return self._value


class DateCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: DateTime) -> None:
        super().__init__(row, col, value)
        self._type = CellType.DATE

    @property
    def value(self) -> datetime:
        return self._value


class DurationCell(Cell):
    def __init__(self, row: int, col: int, value: Duration) -> None:
        super().__init__(row, col, value)
        self._type = CellType.DURATION

    @property
    def value(self) -> duration:
        return self._value


class ErrorCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int) -> None:
        super().__init__(row, col, None)
        self._type = CellType.ERROR

    @property
    def value(self):
        return None


class MergedCell(Cell):
    """.. NOTE::
    Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int) -> None:
        super().__init__(row, col, None)
        self._type = CellType.MERGED

    @property
    def value(self):
        return None


def _pack_decimal128(value: float) -> bytearray:
    buffer = bytearray(16)
    exp = math.floor(math.log10(math.e) * math.log(abs(value))) if value != 0.0 else 0
    exp += 0x1820 - 16
    mantissa = abs(int(value / math.pow(10, exp - 0x1820)))
    buffer[15] |= exp >> 7
    buffer[14] |= (exp & 0x7F) << 1
    i = 0
    while mantissa >= 1:
        buffer[i] = mantissa & 0xFF
        i += 1
        mantissa = int(mantissa / 256)
    if value < 0:
        buffer[15] |= 0x80
    return buffer


def _unpack_decimal128(buffer: bytearray) -> float:
    exp = (((buffer[15] & 0x7F) << 7) | (buffer[14] >> 1)) - 0x1820
    mantissa = buffer[14] & 1
    for i in range(13, -1, -1):
        mantissa = mantissa * 256 + buffer[i]
    sign = 1 if buffer[15] & 0x80 else 0
    if sign == 1:
        mantissa = -mantissa
    value = mantissa * 10**exp
    return float(value)


def _decode_date_format_field(field: str, value: datetime) -> str:
    if field in DATETIME_FIELD_MAP:
        s = DATETIME_FIELD_MAP[field]
        if callable(s):
            return s(value)
        else:
            return value.strftime(s)
    else:
        warn(f"Unsupported field code '{field}'", UnsupportedWarning, stacklevel=4)
        return ""


def _decode_date_format(format, value):
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
                    result += _decode_date_format_field(field, value)
                    in_field = False
                index += 1
        elif in_string:
            result += current_char
            index += 1
        elif not current_char.isalpha():
            if in_field:
                result += _decode_date_format_field(field, value)
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
        result += _decode_date_format_field(field, value)

    return result


def _decode_text_format(format, value: str):
    """Parse a custom date format string and return a formatted number value."""
    custom_format_string = format.custom_format_string
    return custom_format_string.replace(CUSTOM_TEXT_PLACEHOLDER, value)


def _expand_quotes(value: str) -> str:
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


def _decode_number_format(format, value, name):  # noqa: PLR0912
    """Parse a custom date format string and return a formatted number value."""
    custom_format_string = format.custom_format_string
    value *= format.scale_factor
    if "%" in custom_format_string and format.scale_factor == 1.0:
        # Per cent scale has 100x but % does not
        value *= 100.0

    if format.currency_code != "":
        # Replace currency code with symbol and no-break space
        custom_format_string = custom_format_string.replace(
            "\u00a4",
            format.currency_code + "\u00a0",
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
        return _expand_quotes(formatted_value)

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
    return _expand_quotes(formatted_value)


def _format_decimal(value: float, format, percent: bool = False) -> str:
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
                formatted_value,
                decimals=format.decimal_places,
                type=str,
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


def _format_currency(value: float, format) -> str:
    formatted_value = _format_decimal(value, format)
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


INT_TO_BASE_CHAR = [str(x) for x in range(10)] + [chr(x) for x in range(ord("A"), ord("Z") + 1)]


def _invert_bit_str(value: str) -> str:
    """Invert a binary value."""
    return "".join(["0" if b == "1" else "1" for b in value])


def _twos_complement(value: int, base: int) -> str:
    """Calculate the twos complement of a negative integer with minimum 32-bit precision."""
    num_bits = max([32, math.ceil(math.log2(abs(value))) + 1])
    bin_value = bin(abs(value))[2:]
    inverted_bin_value = _invert_bit_str(bin_value).rjust(num_bits, "1")
    twos_complement_dec = int(inverted_bin_value, 2) + 1

    if base == 2:
        return bin(twos_complement_dec)[2:].rjust(num_bits, "1")
    elif base == 8:
        return oct(twos_complement_dec)[2:]
    else:
        return hex(twos_complement_dec)[2:].upper()


def _format_base(value: float, format) -> str:
    if value == 0:
        return "0".zfill(format.base_places)

    value = round(value)

    is_negative = False
    if not format.base_use_minus_sign and format.base in [2, 8, 16]:
        if value < 0:
            return _twos_complement(value, format.base)
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


def _format_fraction_parts_to(whole: int, numerator: int, denominator: int):
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


def _float_to_fraction(value: float, denominator: int) -> str:
    """Convert a float to the nearest fraction and return as a string."""
    whole = int(value)
    numerator = round(denominator * (value - whole))
    return _format_fraction_parts_to(whole, numerator, denominator)


def _float_to_n_digit_fraction(value: float, max_digits: int) -> str:
    """Convert a float to a fraction of a maxinum number of digits
    and return as a string.
    """
    max_denominator = 10**max_digits - 1
    (numerator, denominator) = (
        Fraction.from_float(value).limit_denominator(max_denominator).as_integer_ratio()
    )
    whole = int(value)
    numerator -= whole * denominator
    return _format_fraction_parts_to(whole, numerator, denominator)


def _format_fraction(value: float, format) -> str:
    accuracy = format.fraction_accuracy
    if accuracy & 0xFF000000:
        num_digits = 0x100000000 - accuracy
        return _float_to_n_digit_fraction(value, num_digits)
    else:
        return _float_to_fraction(value, accuracy)


def _format_scientific(value: float, format) -> str:
    formatted_value = sigfig.round(value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
    return f"{formatted_value:.{format.decimal_places}E}"


def _unit_format(unit: str, value: int, style: int, abbrev: Optional[str] = None):
    plural = "" if value == 1 else "s"
    if abbrev is None:
        abbrev = unit[0]
    if style == DurationStyle.COMPACT:
        return ""
    elif style == DurationStyle.SHORT:
        return f"{abbrev}"
    else:
        return f" {unit}" + plural


def _auto_units(cell_value, format):
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


# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str: str) -> tuple:
    """Convert a cell reference in A1 notation to a zero indexed row and column.

    Parameters
    ----------
    cell_str:  str
        A1 notation cell reference

    Returns
    -------
    row, col: int, int
        Cell row and column numbers (zero indexed).
    """
    if not cell_str:
        return 0, 0

    match = range_parts.match(cell_str)
    if not match:
        msg = f"invalid cell reference {cell_str}"
        raise IndexError(msg)

    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord("A") + 1) * (26**expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col


def xl_range(first_row, first_col, last_row, last_col):
    """Convert zero indexed row and col cell references to a A1:B1 range string.

    Parameters
    ----------
    first_row: int
        The first cell row.
    first_col: int
        The first cell column.
    last_row: int
        The last cell row.
    last_col: int
        The last cell column.

    Returns
    -------
    str:
        A1:B1 style range string.
    """
    range1 = xl_rowcol_to_cell(first_row, first_col)
    range2 = xl_rowcol_to_cell(last_row, last_col)

    if range1 == range2:
        return range1
    else:
        return range1 + ":" + range2


def xl_rowcol_to_cell(row, col, row_abs=False, col_abs=False):
    """Convert a zero indexed row and column cell reference to a A1 style string.

    Parameters
    ----------
    row: int
         The cell row.
    col: int
        The cell column.
    row_abs: bool
        If ``True``, make the row absolute.
    col_abs: bool
        If ``True``, make the column absolute.

    Returns
    -------
    str:
        A1 style string.
    """
    if row < 0:
        msg = f"row reference {row} below zero"
        raise IndexError(msg)

    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    row += 1  # Change to 1-index.
    row_abs = "$" if row_abs else ""

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)


def xl_col_to_name(col, col_abs=False):
    """Convert a zero indexed column cell reference to a string.

    Parameters
    ----------
    col: int
        The column number (zero indexed).
    col_abs: bool, default: False
        If ``True``, make the column absolute.

    Returns
    -------
        str:
            Column in A1 notation.
    """
    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    col += 1  # Change to 1-index.
    col_str = ""
    col_abs = "$" if col_abs else ""

    while col:
        # Set remainder from 1 .. 26
        remainder = col % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col = int((col - 1) / 26)

    return col_abs + col_str


@dataclass()
class Formatting:
    allow_none: bool = False
    base_places: int = 0
    base_use_minus_sign: bool = True
    base: int = 10
    control_format: ControlFormattingType = ControlFormattingType.NUMBER
    currency_code: str = "GBP"
    date_time_format: str = DEFAULT_DATETIME_FORMAT
    decimal_places: int = None
    fraction_accuracy: FractionAccuracy = FractionAccuracy.THREE
    increment: float = 1.0
    maximum: float = 100.0
    minimum: float = 1.0
    popup_values: List[str] = field(default_factory=lambda: ["Item 1"])
    negative_style: NegativeNumberStyle = NegativeNumberStyle.MINUS
    show_thousands_separator: bool = False
    type: FormattingType = FormattingType.NUMBER
    use_accounting_style: bool = False
    _format_id = None

    def __post_init__(self):
        if not isinstance(self.type, FormattingType):
            type_name = type(self.type).__name__
            msg = f"Invalid format type '{type_name}'"
            raise TypeError(msg)

        if self.use_accounting_style and self.negative_style != NegativeNumberStyle.MINUS:
            warn(
                "use_accounting_style overriding negative_style",
                RuntimeWarning,
                stacklevel=4,
            )

        if self.type == FormattingType.DATETIME:
            formats = re.sub(r"[^a-zA-Z\s]", " ", self.date_time_format).split()
            for el in formats:
                if el not in DATETIME_FIELD_MAP:
                    msg = f"Invalid format specifier '{el}' in date/time format"
                    raise TypeError(msg)

        if self.type == FormattingType.CURRENCY and self.currency_code not in CURRENCIES:
            raise TypeError(f"Unsupported currency code '{self.currency_code}'")

        if self.decimal_places is None:
            if self.type == FormattingType.CURRENCY:
                self.decimal_places = 2
            else:
                self.decimal_places = DECIMAL_PLACES_AUTO

        if (
            self.type == FormattingType.BASE
            and not self.base_use_minus_sign
            and self.base not in (2, 8, 16)
        ):
            msg = f"base_use_minus_sign must be True for base {self.base}"
            raise TypeError(msg)

        if self.type == FormattingType.BASE and (self.base < 2 or self.base > MAX_BASE):
            msg = "base must be in range 2-36"
            raise TypeError(msg)


@dataclass
class CustomFormatting:
    type: CustomFormattingType = CustomFormattingType.NUMBER
    name: str = None
    integer_format: PaddingType = PaddingType.NONE
    decimal_format: PaddingType = PaddingType.NONE
    num_integers: int = 0
    num_decimals: int = 0
    show_thousands_separator: bool = False
    format: str = "%s"

    def __post_init__(self):
        if not isinstance(self.type, CustomFormattingType):
            type_name = type(self.type).__name__
            msg = f"Invalid format type '{type_name}'"
            raise TypeError(msg)

        if self.type == CustomFormattingType.TEXT and self.format.count("%s") > 1:
            raise TypeError("Custom formats only allow one text substitution")

    @classmethod
    def from_archive(cls, archive: object):
        if archive.format_type == FormatType.CUSTOM_DATE:
            format_type = CustomFormattingType.DATETIME
        elif archive.format_type == FormatType.CUSTOM_NUMBER:
            format_type = CustomFormattingType.NUMBER
        else:
            format_type = CustomFormattingType.TEXT

        return CustomFormatting(name=archive.name, type=format_type)
