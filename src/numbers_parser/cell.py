import re
from collections import namedtuple
from dataclasses import dataclass
from datetime import datetime as builtin_datetime
from datetime import timedelta as builtin_timedelta
from enum import IntEnum
from typing import Any, List, Tuple, Union
from warnings import warn

import sigfig
from pendulum import DateTime, Duration, duration
from pendulum import instance as pendulum_instance

from numbers_parser.cell_storage import CellStorage, CellType
from numbers_parser.constants import (
    DEFAULT_ALIGNMENT,
    DEFAULT_BORDER_COLOR,
    DEFAULT_BORDER_STYLE,
    DEFAULT_BORDER_WIDTH,
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_TEXT_INSET,
    DEFAULT_TEXT_WRAP,
    EMPTY_STORAGE_BUFFER,
    MAX_SIGNIFICANT_DIGITS,
)
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated.TSWPArchives_pb2 import (
    ParagraphStylePropertiesArchive as ParagraphStyle,
)
from numbers_parser.numbers_cache import Cacheable, cache

__all__ = [
    "Alignment",
    "BackgroundImage",
    "BoolCell",
    "Border",
    "BorderType",
    "BulletedTextCell",
    "Cell",
    "CellBorder",
    "DateCell",
    "DurationCell",
    "EmptyCell",
    "ErrorCell",
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
    def __init__(self, image_data: bytes = None, filename: str = None):
        self._data = image_data
        self._filename = filename

    @property
    def data(self) -> bytes:
        """The background image as byts for a cell, or None if no image."""
        return self._data

    @property
    def filename(self) -> str:
        """The image filename for a cell, or None if no image."""
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
                raise TypeError("invalid horizontal alignment")
            horizontal = HORIZONTAL_MAP[horizontal]

        if isinstance(vertical, str):
            vertical = vertical.lower()
            if vertical not in VERTICAL_MAP:
                raise TypeError("invalid vertical alignment")
            vertical = VERTICAL_MAP[vertical]

        return super(_Alignment, cls).__new__(cls, (horizontal, vertical))


DEFAULT_ALIGNMENT_CLASS = Alignment(*DEFAULT_ALIGNMENT)

RGB = namedtuple("RGB", ["r", "g", "b"])


@dataclass
class Style:
    alignment: Alignment = DEFAULT_ALIGNMENT_CLASS
    bg_image: object = None
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
            "first_indent",
            "left_indent",
            "right_indent",
            "text_inset",
            "text_wrap",
        ]

    @classmethod
    def from_storage(cls, cell_storage: object, model: object):
        if cell_storage.image_data is not None:
            bg_image = BackgroundImage(*cell_storage.image_data)
        else:
            bg_image = None
        return Style(
            alignment=model.cell_alignment(cell_storage),
            bg_image=bg_image,
            bg_color=model.cell_bg_color(cell_storage),
            font_color=model.cell_font_color(cell_storage),
            font_size=model.cell_font_size(cell_storage),
            font_name=model.cell_font_name(cell_storage),
            bold=model.cell_is_bold(cell_storage),
            italic=model.cell_is_italic(cell_storage),
            strikethrough=model.cell_is_strikethrough(cell_storage),
            underline=model.cell_is_underline(cell_storage),
            name=model.cell_style_name(cell_storage),
            first_indent=model.cell_first_indent(cell_storage),
            left_indent=model.cell_left_indent(cell_storage),
            right_indent=model.cell_right_indent(cell_storage),
            text_inset=model.cell_text_inset(cell_storage),
            text_wrap=model.cell_text_wrap(cell_storage),
            _text_style_obj_id=model.text_style_object_id(cell_storage),
            _cell_style_obj_id=model.cell_style_object_id(cell_storage),
        )

    def __post_init__(self):
        self.bg_color = rgb_color(self.bg_color)
        self.font_color = rgb_color(self.font_color)

        if not isinstance(self.font_size, float):
            raise TypeError("size must be a float number of points")
        if not isinstance(self.font_name, str):
            raise TypeError("font name must be a string")

        for field in ["bold", "italic", "underline", "strikethrough"]:
            if not isinstance(getattr(self, field), bool):
                raise TypeError(f"{field} argument must be boolean")

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
            raise TypeError("RGB color must be an RGB or a tuple of 3 integers")
        return RGB(*color)
    elif isinstance(color, list):
        return [rgb_color(c) for c in color]
    raise TypeError("RGB color must be an RGB or a tuple of 3 integers")


def alignment(value) -> Alignment:
    """Raise a TypeError if a alignment is not a valid."""
    if value is None:
        return Alignment()
    if isinstance(value, Alignment):
        return value
    if isinstance(value, tuple):
        if not (len(value) == 2 and all(isinstance(x, (int, str)) for x in value)):
            raise TypeError("Alignment must be an Alignment or a tuple of 2 integers/strings")
        return Alignment(*value)
    raise TypeError("Alignment must be an Alignment or a tuple of 2 integers/strings")


BORDER_STYLE_MAP = {"solid": 0, "dashes": 1, "dots": 2, "none": 3}


class BorderType(IntEnum):
    SOLID = BORDER_STYLE_MAP["solid"]
    DASHES = BORDER_STYLE_MAP["dashes"]
    DOTS = BORDER_STYLE_MAP["dots"]
    NONE = BORDER_STYLE_MAP["none"]


class Border:
    def __init__(
        self,
        width: float = DEFAULT_BORDER_WIDTH,
        color: RGB = None,
        style: BorderType = None,
        _order: int = 0,
    ):
        if not isinstance(width, float):
            raise TypeError("width must be a float number of points")
        self.width = width

        if color is None:
            color = RGB(*DEFAULT_BORDER_COLOR)
        self.color = rgb_color(color)

        if style is None:
            style = BorderType(BORDER_STYLE_MAP[DEFAULT_BORDER_STYLE])
        if isinstance(style, str):
            style = style.lower()
            if style not in BORDER_STYLE_MAP:
                raise TypeError("invalid border style")
            self.style = BORDER_STYLE_MAP[style]
        else:
            self.style = style

        self._order = _order

    def __repr__(self) -> str:
        style_name = BorderType(self.style).name.lower()
        return f"Border(width={self.width}, color={self.color}, style={style_name})"

    def __eq__(self, value: object) -> bool:
        return all(
            [self.width == value.width, self.color == value.color, self.style == value.style]
        )


class CellBorder:
    def __init__(
        self,
        top_merged: bool = False,
        right_merged: bool = False,
        bottom_merged: bool = False,
        left_merged: bool = False,
    ):
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

    def __init__(self, row_start: int, col_start: int, row_end: int, col_end: int):
        self.rect = (row_start, col_start, row_end, col_end)


class MergeAnchor:
    """Cell reference for the merged cell."""

    def __init__(self, size: Tuple):
        self.size = size


class Cell(Cacheable):
    @classmethod
    def empty_cell(cls, table_id: int, row_num: int, col_num: int, model: object):
        cell = EmptyCell(row_num, col_num)
        cell._model = model
        cell._table_id = table_id
        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row_num, col_num)))

        return cell

    @classmethod
    def merged_cell(cls, table_id: int, row_num: int, col_num: int, model: object):
        cell = MergedCell(row_num, col_num)
        cell._model = model
        cell._table_id = table_id
        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row_num, col_num)))
        return cell

    @classmethod
    def from_storage(cls, cell_storage: CellStorage):
        if cell_storage.type == CellType.EMPTY:
            cell = EmptyCell(cell_storage.row_num, cell_storage.col_num)
        elif cell_storage.type == CellType.NUMBER:
            cell = NumberCell(cell_storage.row_num, cell_storage.col_num, cell_storage.value)
        elif cell_storage.type == CellType.TEXT:
            cell = TextCell(cell_storage.row_num, cell_storage.col_num, cell_storage.value)
        elif cell_storage.type == CellType.DATE:
            cell = DateCell(cell_storage.row_num, cell_storage.col_num, cell_storage.value)
        elif cell_storage.type == CellType.BOOL:
            cell = BoolCell(cell_storage.row_num, cell_storage.col_num, cell_storage.value)
        elif cell_storage.type == CellType.DURATION:
            value = duration(seconds=cell_storage.value)
            cell = DurationCell(cell_storage.row_num, cell_storage.col_num, value)
        elif cell_storage.type == CellType.ERROR:
            cell = ErrorCell(cell_storage.row_num, cell_storage.col_num)
        elif cell_storage.type == CellType.RICH_TEXT:
            cell = RichTextCell(cell_storage.row_num, cell_storage.col_num, cell_storage.value)
        else:
            raise UnsupportedError(
                f"Unsupported cell type {cell_storage.type} "
                + f"@:({cell_storage.row_num},{cell_storage.col_num})"
            )

        cell._table_id = cell_storage.table_id
        cell._model = cell_storage.model
        cell._storage = cell_storage
        cell._formula_key = cell_storage.formula_id
        merge_cells = cell_storage.model.merge_cells(cell_storage.table_id)
        cell._set_merge(merge_cells.get((cell_storage.row_num, cell_storage.col_num)))
        return cell

    @classmethod
    def from_value(cls, row_num: int, col_num: int, value):
        # TODO: write needs to retain/init the border
        if isinstance(value, str):
            return TextCell(row_num, col_num, value)
        elif isinstance(value, bool):
            return BoolCell(row_num, col_num, value)
        elif isinstance(value, int):
            return NumberCell(row_num, col_num, value)
        elif isinstance(value, float):
            rounded_value = sigfig.round(value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
            if rounded_value != value:
                warn(
                    f"'{value}' rounded to {MAX_SIGNIFICANT_DIGITS} significant digits",
                    RuntimeWarning,
                    stacklevel=2,
                )
            return NumberCell(row_num, col_num, rounded_value)
        elif isinstance(value, (DateTime, builtin_datetime)):
            return DateCell(row_num, col_num, pendulum_instance(value))
        elif isinstance(value, (Duration, builtin_timedelta)):
            return DurationCell(row_num, col_num, value)
        else:
            raise ValueError("Can't determine cell type from type " + type(value).__name__)

    def set_formatting(self, formatting: dict):
        raise TypeError(f"Cannot set formatting for cells of type {type(self).__name__}")

    def __init__(self, row_num: int, col_num: int, value):
        self._value = value
        self.row = row_num
        self.col = col_num
        self.is_bulleted = False
        self._formula_key = None
        self._storage = None
        self._style = None

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
    def is_formula(self):
        table_formulas = self._model.table_formulas(self._table_id)
        return table_formulas.is_formula(self.row, self.col)

    @property
    @cache(num_args=0)
    def formula(self):
        if self._formula_key is not None:
            table_formulas = self._model.table_formulas(self._table_id)
            return table_formulas.formula(self._formula_key, self.row, self.col)
        else:
            return None

    @property
    def bullets(self) -> str:
        return None

    @property
    def formatted_value(self):
        if self._storage is None:
            return ""
        else:
            return self._storage.formatted

    @property
    def style(self):
        if self._storage is None:
            self._storage = CellStorage(
                self._model, self._table_id, EMPTY_STORAGE_BUFFER, self.row, self.col
            )
        if self._style is None:
            self._style = Style.from_storage(self._storage, self._model)
        return self._style

    @style.setter
    def style(self, _):
        warn(
            "cell style cannot be set; use Table.set_cell_style() instead",
            UnsupportedWarning,
            stacklevel=2,
        )

    @property
    def border(self):
        self._model.extract_strokes(self._table_id)
        return self._border

    @border.setter
    def border(self, _):
        warn(
            "cell border values cannot be set; use Table.set_cell_border() instead",
            UnsupportedWarning,
            stacklevel=2,
        )

    def update_storage(self, storage: CellStorage) -> None:
        self._storage = storage


class NumberCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: float):
        self._type = TSTArchives.numberCellType
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> int:
        return self._value

    def set_formatting(self, formatting: dict):
        self._storage.set_number_formatting(formatting)


class TextCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: str):
        self._type = TSTArchives.textCellType
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> str:
        return self._value


class RichTextCell(Cell):
    def __init__(self, row_num: int, col_num: int, value):
        self._type = TSTArchives.automaticCellType
        super().__init__(row_num, col_num, value["text"])
        self._bullets = value["bullets"]
        self._hyperlinks = value["hyperlinks"]
        if value["bulleted"]:
            self._formatted_bullets = [
                value["bullet_chars"][i] + " " + value["bullets"][i]
                if value["bullet_chars"][i] is not None
                else value["bullets"][i]
                for i in range(len(self._bullets))
            ]
            self.is_bulleted = True

    @property
    def value(self) -> str:
        return self._value

    @property
    def bullets(self) -> str:
        return self._bullets

    @property
    def formatted_bullets(self) -> str:
        return self._formatted_bullets

    @property
    def hyperlinks(self) -> List[Tuple]:
        return self._hyperlinks


# Backwards compatibility to earlier class names
class BulletedTextCell(RichTextCell):
    pass


class EmptyCell(Cell):
    def __init__(self, row_num: int, col_num: int):
        super().__init__(row_num, col_num, None)
        self._type = None

    @property
    def value(self):
        return None


class BoolCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: bool):
        self._type = TSTArchives.boolCellType
        self._value = value
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> bool:
        return self._value


class DateCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: DateTime):
        self._type = TSTArchives.dateCellType
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> duration:
        return self._value

    def set_formatting(self, formatting: dict):
        if "date_time_format" not in formatting:
            raise TypeError("No date_time_format specified for DateCell formatting")
        self._storage.set_date_time_formatting(formatting["date_time_format"])


class DurationCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: Duration):
        self._type = TSTArchives.durationCellType
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> duration:
        return self._value


class ErrorCell(Cell):
    def __init__(self, row_num: int, col_num: int):
        self._type = TSTArchives.formulaErrorCellType
        super().__init__(row_num, col_num, None)

    @property
    def value(self):
        return None


class MergedCell(Cell):
    def __init__(self, row_num: int, col_num: int):
        super().__init__(row_num, col_num, None)

    @property
    def value(self):
        return None


# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str: str) -> tuple:
    """Convert a cell reference in A1 notation to a zero indexed row and column.
    Args:
        cell_str:  A1 style string.
    Returns:
        row, col: Zero indexed cell row and column indices.
    """
    if not cell_str:
        return 0, 0

    match = range_parts.match(cell_str)
    if not match:
        raise IndexError(f"invalid cell reference {cell_str}")

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
    Args:
       first_row: The first cell row.    Int.
       first_col: The first cell column. Int.
       last_row:  The last cell row.     Int.
       last_col:  The last cell column.  Int.
    Returns:
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
    Args:
       row:     The cell row.    Int.
       col:     The cell column. Int.
       row_abs: Optional flag to make the row absolute.    Bool.
       col_abs: Optional flag to make the column absolute. Bool.
    Returns:
        A1 style string.
    """
    if row < 0:
        raise IndexError(f"row reference {row} below zero")

    if col < 0:
        raise IndexError(f"column reference {col} below zero")

    row += 1  # Change to 1-index.
    row_abs = "$" if row_abs else ""

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)


def xl_col_to_name(col, col_abs=False):
    """Convert a zero indexed column cell reference to a string.
    Args:
       col:     The cell column. Int.
       col_abs: Optional flag to make the column absolute. Bool.
    Returns:
        Column style string.
    """
    col_num = col
    if col_num < 0:
        raise IndexError(f"column reference {col_num} below zero")

    col_num += 1  # Change to 1-index.
    col_str = ""
    col_abs = "$" if col_abs else ""

    while col_num:
        # Set remainder from 1 .. 26
        remainder = col_num % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_abs + col_str
