import re
import warnings

from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.generated.TSWPArchives_pb2 import (
    ParagraphStylePropertiesArchive as ParagraphStyle,
)
from numbers_parser.exceptions import UnsupportedError
from numbers_parser.cell_storage import CellType, CellStorage
from numbers_parser.constants import (
    EMPTY_STORAGE_BUFFER,
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_ALIGNMENT,
)

from dataclasses import dataclass
from collections import namedtuple
from enum import IntEnum
from functools import lru_cache
from pendulum import duration, Duration, DateTime
from typing import Any, List, Tuple, Union


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


# Style return types
_Alignment = namedtuple("Alignment", ["horizontal", "vertical"])
RGB = namedtuple("RGB", ["r", "g", "b"])


class HorizontalJustification(IntEnum):
    LEFT = ParagraphStyle.TextAlignmentType.TATvalue0
    RIGHT = ParagraphStyle.TextAlignmentType.TATvalue1
    CENTER = ParagraphStyle.TextAlignmentType.TATvalue2
    JUSTIFIED = ParagraphStyle.TextAlignmentType.TATvalue3
    AUTO = ParagraphStyle.TextAlignmentType.TATvalue4


class VerticalJustification(IntEnum):
    TOP = ParagraphStyle.ParagraphBorderType.PBTvalue0
    MIDDLE = ParagraphStyle.ParagraphBorderType.PBTvalue1
    BOTTOM = ParagraphStyle.ParagraphBorderType.PBTvalue2


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

        self = super(_Alignment, cls).__new__(cls, (horizontal, vertical))
        return self


DEFAULT_ALIGNMENT_CLASS = Alignment(*DEFAULT_ALIGNMENT)


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
    name: str = None
    _text_style_obj_id: int = None
    _cell_style_obj_id: int = None
    _update_cell_style: bool = False
    _update_text_style: bool = False

    @staticmethod
    def _text_attrs():
        return [
            "font_color",
            "font_size",
            "font_name",
            "bold",
            "italic",
            "strikethrough",
            "underline",
            "name",
        ]

    @classmethod
    def from_storage(cls, cell_storage: object, model: object):
        style = Style()

        if cell_storage.image_data is not None:
            bg_image = BackgroundImage(*cell_storage.image_data)
        else:
            bg_image = None
        style = Style(
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
            _text_style_obj_id=model.text_style_object_id(cell_storage),
            _cell_style_obj_id=model.cell_style_object_id(cell_storage),
        )
        return style

    def __post_init__(self):
        if not isinstance(self.alignment, Alignment):
            raise TypeError("value must be an Alignment class")

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
        possible updates when saving the document"""
        if name in self.__dict__:
            # Init has been done
            if name in ["bg_color", "font_color"]:
                value = rgb_color(value)
            if name in ["bg_color"]:
                self.__dict__["_update_cell_style"] = True
            elif name in ["alignment"]:
                self.__dict__["_update_text_style"] = True
                self.__dict__["_update_cell_style"] = True
            elif name in Style._text_attrs():
                self.__dict__["_update_text_style"] = True
        elif name == "bg_color" and value is not None:
            self.__dict__["_update_cell_style"] = True
        elif (
            name == "alignment"
            and value is not None
            and value[1] != DEFAULT_ALIGNMENT_CLASS[1]
        ):
            self.__dict__["_update_cell_style"] = True
        if name not in ["_update_text_style", "_update_cell_style"]:
            self.__dict__[name] = value


def rgb_color(color) -> RGB:
    """Raise a TypeError if a color is not a valid RGB value"""
    if color is None:
        return None
    if isinstance(color, RGB):
        return color
    if isinstance(color, tuple):
        if not (len(color) == 3 and all([isinstance(x, int) for x in color])):
            raise TypeError("RGB color must be an RGB or a tuple of 3 integers")
        return RGB(*color)
    elif isinstance(color, list):
        return [rgb_color(c) for c in color]


class Cell:
    @classmethod
    def empty_cell(cls, table_id: int, row_num: int, col_num: int, model: object):
        row_col = (row_num, col_num)
        merge_cells = model.merge_cell_ranges(table_id)
        is_merged = row_col in merge_cells

        if is_merged and merge_cells[row_col]["merge_type"] == "ref":
            cell = MergedCell(*merge_cells[row_col]["rect"])
        else:
            cell = EmptyCell(row_num, col_num)
        cell.size = None
        cell._model = model
        cell._table_id = table_id
        cell.style = None
        return cell

    @classmethod
    def from_storage(cls, cell_storage: CellStorage):  # noqa: C901
        row_col = (cell_storage.row_num, cell_storage.col_num)
        merge_cells = cell_storage.model.merge_cell_ranges(cell_storage.table_id)
        is_merged = row_col in merge_cells

        if cell_storage.type == CellType.EMPTY:
            cell = EmptyCell(cell_storage.row_num, cell_storage.col_num)
        elif cell_storage.type == CellType.NUMBER:
            cell = NumberCell(*row_col, cell_storage.value)
        elif cell_storage.type == CellType.TEXT:
            cell = TextCell(*row_col, cell_storage.value)
        elif cell_storage.type == CellType.DATE:
            cell = DateCell(*row_col, cell_storage.value)
        elif cell_storage.type == CellType.BOOL:
            cell = BoolCell(*row_col, cell_storage.value)
        elif cell_storage.type == CellType.DURATION:
            value = duration(seconds=cell_storage.value)
            cell = DurationCell(*row_col, value)
        elif cell_storage.type == CellType.ERROR:
            cell = ErrorCell(*row_col)
        elif cell_storage.type == CellType.RICH_TEXT:
            cell = RichTextCell(*row_col, cell_storage.value)
        else:
            raise UnsupportedError(
                f"Unsupported cell type {cell_storage.type} "
                + f"@:({cell_storage.row_num},{cell_storage.col_num})"
            )

        cell._table_id = cell_storage.table_id
        cell._model = cell_storage.model
        cell._storage = cell_storage
        cell._formula_key = cell_storage.formula_id
        cell._style = None

        if is_merged and merge_cells[row_col]["merge_type"] == "source":
            cell.is_merged = True
            cell.size = merge_cells[row_col]["size"]

        return cell

    def __init__(self, row_num: int, col_num: int, value):
        self._value = value
        self.row = row_num
        self.col = col_num
        self.size = (1, 1)
        self.is_merged = False
        self.is_bulleted = False
        self._formula_key = None
        self._storage = None
        self._style = None

    @property
    def image_filename(self):
        warnings.warn(
            "image_filename is deprecated and will be removed in the future. "
            + "Please use the style property",
            DeprecationWarning,
        )
        if self.style is not None and self.style.bg_image is not None:
            return self.style.bg_image.filename
        else:
            return None

    @property
    def image_data(self):
        warnings.warn(
            "image_data is deprecated and will be removed in the future. "
            + "Please use the style property",
            DeprecationWarning,
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
    @lru_cache(maxsize=None)
    def formula(self):
        if self._formula_key is not None:
            table_formulas = self._model.table_formulas(self._table_id)
            formula = table_formulas.formula(self._formula_key, self.row, self.col)
            return formula
        else:
            return None

    @property
    def bullets(self) -> str:
        return None

    @property
    def formatted_value(self):
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
    def style(self, style):
        self._style = style


class NumberCell(Cell):
    def __init__(self, row_num: int, col_num: int, value: float):
        self._type = TSTArchives.numberCellType
        super().__init__(row_num, col_num, value)

    @property
    def value(self) -> int:
        return self._value


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
        self._type = None
        super().__init__(row_num, col_num, None)

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


class MergedCell:
    def __init__(self, row_start: int, col_start: int, row_end: int, col_end: int):
        self.value = None
        self.formula = None
        self.row_start = row_start
        self.row_end = row_end
        self.col_start = col_start
        self.col_end = col_end
        self.merge_range = xl_range(row_start, col_start, row_end, col_end)


# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str: str) -> tuple:
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.
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
    """
    Convert zero indexed row and col cell references to a A1:B1 range string.
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
    """
    Convert a zero indexed row and column cell reference to a A1 style string.
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
    """
    Convert a zero indexed column cell reference to a string.
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
