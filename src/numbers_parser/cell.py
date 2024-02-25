import re
from collections import namedtuple
from dataclasses import dataclass, field
from datetime import datetime as builtin_datetime
from datetime import timedelta as builtin_timedelta
from enum import IntEnum
from os.path import basename
from typing import Any, List, Tuple, Union
from warnings import warn

import sigfig
from pendulum import DateTime, Duration, duration
from pendulum import instance as pendulum_instance

from numbers_parser.cell_storage import CellStorage, CellType
from numbers_parser.constants import (
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
    MAX_BASE,
    MAX_SIGNIFICANT_DIGITS,
    ControlFormattingType,
    CustomFormattingType,
    FormattingType,
    FormatType,
    FractionAccuracy,
    NegativeNumberStyle,
    PaddingType,
)
from numbers_parser.currencies import CURRENCIES
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
    """
    A named document style that can be applied to cells.

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

    def __init__(self, data: bytes = None, filename: str = None):
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
    """
    A named document style that can be applied to cells.

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

        for attr in ["bold", "italic", "underline", "strikethrough"]:
            if not isinstance(getattr(self, attr), bool):
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
    """
    Create a cell border to use with the :py:class:`~numbers_parser.Table` method
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
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    @classmethod
    def empty_cell(cls, table_id: int, row: int, col: int, model: object):
        cell = EmptyCell(row, col)
        cell._model = model
        cell._table_id = table_id
        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row, col)))

        return cell

    @classmethod
    def merged_cell(cls, table_id: int, row: int, col: int, model: object):
        cell = MergedCell(row, col)
        cell._model = model
        cell._table_id = table_id
        merge_cells = model.merge_cells(table_id)
        cell._set_merge(merge_cells.get((row, col)))
        return cell

    @classmethod
    def from_storage(cls, cell_storage: CellStorage):
        if cell_storage.type == CellType.EMPTY:
            cell = EmptyCell(cell_storage.row, cell_storage.col)
        elif cell_storage.type == CellType.NUMBER:
            cell = NumberCell(cell_storage.row, cell_storage.col, cell_storage.value)
        elif cell_storage.type == CellType.TEXT:
            cell = TextCell(cell_storage.row, cell_storage.col, cell_storage.value)
        elif cell_storage.type == CellType.DATE:
            cell = DateCell(cell_storage.row, cell_storage.col, cell_storage.value)
        elif cell_storage.type == CellType.BOOL:
            cell = BoolCell(cell_storage.row, cell_storage.col, cell_storage.value)
        elif cell_storage.type == CellType.DURATION:
            value = duration(seconds=cell_storage.value)
            cell = DurationCell(cell_storage.row, cell_storage.col, value)
        elif cell_storage.type == CellType.ERROR:
            cell = ErrorCell(cell_storage.row, cell_storage.col)
        elif cell_storage.type == CellType.RICH_TEXT:
            cell = RichTextCell(cell_storage.row, cell_storage.col, cell_storage.value)
        else:
            raise UnsupportedError(
                f"Unsupported cell type {cell_storage.type} "
                + f"@:({cell_storage.row},{cell_storage.col})"
            )

        cell._table_id = cell_storage.table_id
        cell._model = cell_storage.model
        cell._storage = cell_storage
        cell._formula_key = cell_storage.formula_id
        merge_cells = cell_storage.model.merge_cells(cell_storage.table_id)
        cell._set_merge(merge_cells.get((cell_storage.row, cell_storage.col)))
        return cell

    @classmethod
    def from_value(cls, row: int, col: int, value):
        # TODO: write needs to retain/init the border
        if isinstance(value, str):
            return TextCell(row, col, value)
        elif isinstance(value, bool):
            return BoolCell(row, col, value)
        elif isinstance(value, int):
            return NumberCell(row, col, value)
        elif isinstance(value, float):
            rounded_value = sigfig.round(value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
            if rounded_value != value:
                warn(
                    f"'{value}' rounded to {MAX_SIGNIFICANT_DIGITS} significant digits",
                    RuntimeWarning,
                    stacklevel=2,
                )
            return NumberCell(row, col, rounded_value)
        elif isinstance(value, (DateTime, builtin_datetime)):
            return DateCell(row, col, pendulum_instance(value))
        elif isinstance(value, (Duration, builtin_timedelta)):
            return DurationCell(row, col, value)
        else:
            raise ValueError("Can't determine cell type from type " + type(value).__name__)

    def _set_formatting(
        self,
        format_id: int,
        format_type: FormattingType,
        control_id: int,
        is_currency: bool = False,
    ) -> None:
        self._storage._set_formatting(format_id, format_type, control_id, is_currency)

    def __init__(self, row: int, col: int, value):
        self._value = value
        self.row = row
        self.col = col
        self._is_bulleted = False
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
    def is_formula(self) -> bool:
        """bool: ``True`` if the cell contains a formula."""
        table_formulas = self._model.table_formulas(self._table_id)
        return table_formulas.is_formula(self.row, self.col)

    @property
    @cache(num_args=0)
    def formula(self) -> str:
        """
        str: The formula in a cell.

        Formula evaluation relies on Numbers storing current values which should
        usually be the case. In cells containing a formula, :py:meth:`numbers_parser.Cell.value`
        returns computed value of the formula.

        Returns:
            str:
                The text of the foruma in a cell, or `None` if there is no formula
                present in a cell.
        """
        if self._formula_key is not None:
            table_formulas = self._model.table_formulas(self._table_id)
            return table_formulas.formula(self._formula_key, self.row, self.col)
        else:
            return None

    @property
    def is_bulleted(self) -> bool:
        """bool: ``True`` if the cell contains text bullets."""
        return self._is_bulleted

    @property
    def bullets(self) -> Union[List[str], None]:
        r"""
        List[str] | None: The bullets in a cell, or ``None``

        Cells that contain bulleted or numbered lists are identified
        by :py:attr:`numbers_parser.Cell.is_bulleted`. For these cells,
        :py:attr:`numbers_parser.Cell.value` returns the whole cell contents.
        Bullets can also be extracted into a list of paragraphs cell without the
        bullet or numbering character. Newlines are not included in the
        bullet list.

        Example
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
        """
        str: The formatted value of the cell as it appears in Numbers.

        Interactive elements are converted into a suitable text format where
        supported, or as their number values where there is no suitable
        visual representation. Currently supported mappings are:

        * Checkboxes are U+2610 (Ballow Box) or U+2611 (Ballot Box with Check)
        * Ratings are their star value represented using (U+2605) (Black Star)

        .. code-block:: python

            >>> table = doc.sheets[0].tables[0]
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
        if self._storage is None:
            return ""
        else:
            return self._storage.formatted

    @property
    def style(self) -> Union[Style, None]:
        """Style | None: The :class:`Style` associated with the cell or ``None``.

        Warns:
            UnsupportedWarning: On assignment; use
                :py:meth:`numbers_parser.Table.set_cell_style` instead.
        """
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
    def border(self) -> Union[CellBorder, None]:
        """CellBorder| None: The :class:`CellBorder` associated with the cell or ``None``.

        Warns:
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

    def update_storage(self, storage: CellStorage) -> None:
        self._storage = storage


class NumberCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: float):
        self._type = TSTArchives.numberCellType
        super().__init__(row, col, value)

    @property
    def value(self) -> int:
        return self._value


class TextCell(Cell):
    def __init__(self, row: int, col: int, value: str):
        self._type = TSTArchives.textCellType
        super().__init__(row, col, value)

    @property
    def value(self) -> str:
        return self._value


class RichTextCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value):
        self._type = TSTArchives.automaticCellType
        super().__init__(row, col, value["text"])
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
        """
        List[Tuple] | None: the hyperlinks in a cell or ``None``

        Numbers does not support hyperlinks to cells within a spreadsheet, but does
        allow embedding links in cells. When cells contain hyperlinks,
        `numbers_parser` returns the text version of the cell. The `hyperlinks` property
        of cells where :py:attr:`numbers_parser.Cell.is_bulleted` is ``True`` is a
        list of text and URL tuples.

        Example
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
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int):
        super().__init__(row, col, None)
        self._type = None

    @property
    def value(self):
        return None


class BoolCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: bool):
        self._type = TSTArchives.boolCellType
        self._value = value
        super().__init__(row, col, value)

    @property
    def value(self) -> bool:
        return self._value


class DateCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int, value: DateTime):
        self._type = TSTArchives.dateCellType
        super().__init__(row, col, value)

    @property
    def value(self) -> duration:
        return self._value


class DurationCell(Cell):
    def __init__(self, row: int, col: int, value: Duration):
        self._type = TSTArchives.durationCellType
        super().__init__(row, col, value)

    @property
    def value(self) -> duration:
        return self._value


class ErrorCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int):
        self._type = TSTArchives.formulaErrorCellType
        super().__init__(row, col, None)

    @property
    def value(self):
        return None


class MergedCell(Cell):
    """
    .. NOTE::
       Do not instantiate directly. Cells are created by :py:class:`~numbers_parser.Document`.
    """

    def __init__(self, row: int, col: int):
        super().__init__(row, col, None)

    @property
    def value(self):
        return None


# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str: str) -> tuple:
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

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
    """
    Convert a zero indexed row and column cell reference to a A1 style string.

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

    Parameters
    ----------
    col: int
        The column number (zero indexed).
    col_abs: bool, default: False
        If ``True``, make the column absolute.
    Returns:
        str:
            Column in A1 notation.
    """
    if col < 0:
        raise IndexError(f"column reference {col} below zero")

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
    negative_style: NegativeNumberStyle = NegativeNumberStyle.MINUS
    show_thousands_separator: bool = False
    type: FormattingType = FormattingType.NUMBER
    use_accounting_style: bool = False
    values: List[str] = field(default_factory=lambda: ["Item 1"])
    _format_id = None

    def __post_init__(self):
        if not isinstance(self.type, FormattingType):
            type_name = type(self.type).__name__
            raise TypeError(f"Invalid format type '{type_name}'")

        if self.use_accounting_style and self.negative_style != NegativeNumberStyle.MINUS:
            warn(
                "use_accounting_style overriding negative_style",
                RuntimeWarning,
                stacklevel=4,
            )

        if self.type == FormattingType.DATETIME:
            formats = re.sub(r"[^a-zA-Z\s]", " ", self.date_time_format).split()
            for el in formats:
                if el not in DATETIME_FIELD_MAP.keys():
                    raise TypeError(f"Invalid format specifier '{el}' in date/time format")

        if self.type == FormattingType.CURRENCY:
            if self.currency_code not in CURRENCIES:
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
            raise TypeError(f"base_use_minus_sign must be True for base {self.base}")

        if self.type == FormattingType.BASE and (self.base < 2 or self.base > MAX_BASE):
            raise TypeError("base must be in range 2-36")


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
            raise TypeError(f"Invalid format type '{type_name}'")

        if self.type == CustomFormattingType.TEXT:
            if self.format.count("%s") > 1:
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
