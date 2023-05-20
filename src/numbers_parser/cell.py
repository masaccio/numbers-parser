import re

from numbers_parser.generated import TSTArchives_pb2 as TSTArchives
from numbers_parser.exceptions import UnsupportedError
from numbers_parser.cell_storage import CellType, CellStorage

from pendulum import duration, Duration, DateTime
from functools import lru_cache
from typing import List, Tuple


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
            if cell_storage.value is None:
                value = duration(seconds=0)
            else:
                value = duration(seconds=cell_storage.value)
            cell = DurationCell(*row_col, value)
        elif cell_storage.type == CellType.ERROR:
            cell = ErrorCell(*row_col)
        elif cell_storage.type == CellType.RICH_TEXT:
            cell = RichTextCell(*row_col, cell_storage.value)
        else:
            raise UnsupportedError(  # pragma: no cover
                f"Unsupport cell type {cell_storage.type} "
                + "@:({cell_storage.row_num},{cell_storage.col_num})"
            )

        cell._table_id = cell_storage.table_id
        cell._model = cell_storage.model
        cell._storage = cell_storage
        cell._formula_key = cell_storage.formula_id

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
        if self._storage is None:
            return str(self._value)
        else:
            return self._storage.formatted

    @property
    def image_data(self) -> bytes:
        """Return the background image for a cell, or None if no image"""
        data = self._storage.image_data
        if data is None:
            return None
        return data[0]

    @property
    def image_filename(self) -> bytes:
        """Return the background image filename for a cell, or None if no image"""
        data = self._storage.image_data
        if data is None:
            return None
        return self._storage.image_data[1]


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
    if row < 0:  # pragma: no cover
        raise IndexError(f"Row reference {row} below zero")

    if col < 0:  # pragma: no cover
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
    if col_num < 0:  # pragma: no cover
        raise IndexError(f"Column reference {col_num} below zero")

    col_num += 1  # Change to 1-index.
    col_str = ""
    col_abs = "$" if col_abs else ""

    while col_num:
        # Set remainder from 1 .. 26
        remainder = col_num % 26

        if remainder == 0:  # pragma: no cover
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_abs + col_str
