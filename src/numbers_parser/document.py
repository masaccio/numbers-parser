from functools import lru_cache
from logging import debug
from typing import Union, Generator

from numbers_parser.containers import ItemsList
from numbers_parser.model import NumbersModel

from numbers_parser.cell import (
    Cell,
    EmptyCell,
    MergedCell,
    xl_cell_to_rowcol,
    xl_range,
)


class Document:
    def __init__(self, filename):
        self._model = NumbersModel(filename)

    def sheets(self):
        if not hasattr(self, "_sheets"):
            refs = self._model.sheet_ids()
            self._sheets = ItemsList(self._model, refs, Sheet)
        return self._sheets


class Sheet:
    def __init__(self, model, sheet_id):
        self._sheet_id = sheet_id
        self._model = model

    def tables(self):
        if not hasattr(self, "_tables"):
            refs = self._model.table_ids(self._sheet_id)
            self._tables = ItemsList(self._model, refs, Table)
        return self._tables

    @property
    def name(self):
        return self._model.sheet_name(self._sheet_id)


class Table:
    def __init__(self, model, table_id):
        super(Table, self).__init__()
        self._model = model
        self._table_id = table_id

    @property
    def name(self):
        return self._model.table_name(self._table_id)

    @lru_cache(maxsize=None)
    def rows(self, values_only: bool = False) -> list:
        """
        Return all rows of cells for the Table.

        Args:
            values_only: if True, return cell values instead of Cell objects

        Returns:
            rows: list of rows; each row is a list of Cell objects
        """
        row_cells = []
        if values_only:
            for row_num in range(self.num_rows):
                row_cells.append(
                    [
                        self.cell(row_num, col_num).value
                        for col_num in range(self.num_cols)
                    ]
                )
        else:
            for row_num in range(self.num_rows):
                row_cells.append(
                    [self.cell(row_num, col_num) for col_num in range(self.num_cols)]
                )
        return row_cells

    @property
    def merge_ranges(self) -> list:
        if not hasattr(self, "_merge_cells"):
            self._merge_cells = self._model.merge_cell_ranges(self._table_id)

        ranges = [xl_range(*r["rect"]) for r in self._merge_cells.values()]
        return sorted(set(list(ranges)))

    @property
    def _merge_cells(self):
        if not hasattr(self, "__merge_cells"):
            self.__merge_cells = self._model.merge_cell_ranges(self._table_id)
        return self.__merge_cells

    @property
    def num_rows(self) -> int:
        """Number of rows in the table"""
        return len(self._model.table_row_headers(self._table_id))

    @property
    def num_cols(self) -> int:
        """Number of columns in the table"""
        return len(self._model.table_row_columns(self._table_id))

    def cell(self, *args) -> Union[Cell, MergedCell]:
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
        elif len(args) != 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args

        if row_num >= self.num_rows or row_num < 0:
            raise IndexError(f"row {row_num} out of range")
        if col_num >= self.num_cols or col_num < 0:
            raise IndexError(f"coumn {col_num} out of range")

        cell_struct = self._model.table_cell_decode(self._table_id, row_num, col_num)

        row_col = (row_num, col_num)
        is_merged = row_col in self._merge_cells
        if cell_struct is None:
            if is_merged and self._merge_cells[row_col]["merge_type"] == "ref":
                cell = MergedCell(*self._merge_cells[row_col]["rect"])
            else:
                cell = EmptyCell(row_num, col_num, None)
            cell.size = None
            return cell

        cell = Cell.factory(row_num, col_num, cell_struct)
        if is_merged and self._merge_cells[row_col]["merge_type"] == "source":
            cell.is_merged = True
            cell.size = self._merge_cells[row_col]["size"]
        else:
            cell.is_merged = False
            cell.size = (1, 1)

        cell_struct["formula_key"] = self._model.table_cell_formula_decode(
            self._table_id, row_num, col_num, cell_struct["type"]
        )

        if cell_struct["formula_key"] is not None:
            formula = self._model.table_formulas(self._table_id).formula(
                cell_struct["formula_key"], row_num, col_num
            )
            cell.add_formula(formula)

            debug(
                "%s@[%d,%d]: key=%d:%s: type=%d, value=%s",
                self.name,
                row_num,
                col_num,
                cell_struct["formula_key"],
                formula,
                cell_struct["type"],
                str(cell_struct["value"]),
            )
        else:
            debug(
                "%s@[%d,%d]: no formula: type=%d, value=%s",
                self.name,
                row_num,
                col_num,
                cell_struct["type"],
                str(cell_struct["value"]),
            )

        return cell

    def iter_rows(
        self,
        min_row: int = None,
        max_row: int = None,
        min_col: int = None,
        max_col: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """
        Produces cells from a table, by row. Specify the iteration range using
        the indexes of the rows and columns.

        Args:
            min_row: smallest row index (zero indexed)
            max_row: largest row (zero indexed)
            min_col: smallest row index (zero indexed)
            max_col: largest row (zero indexed)
            values_only: return cell values rather than Cell objects

        Returns:
            generator: tuple of cells

        Raises:
            IndexError: row or column values are out of range for the table
        """
        min_row = min_row or 0
        max_row = max_row or self.num_rows - 1
        min_col = min_col or 0
        max_col = max_col or self.num_cols - 1

        if min_row < 0:
            raise IndexError(f"row {min_row} out of range")
        if max_row > self.num_rows:
            raise IndexError(f"row {max_row} out of range")
        if min_col < 0:
            raise IndexError(f"column {min_col} out of range")
        if max_col > self.num_cols:
            raise IndexError(f"column {max_col} out of range")

        rows = self.rows()
        for row_num in range(min_row, max_row + 1):
            if values_only:
                yield tuple(
                    cell.value or None for cell in rows[row_num][min_col : max_col + 1]
                )
            else:
                yield tuple(rows[row_num][min_col : max_col + 1])

    def iter_cols(
        self,
        min_col: int = None,
        max_col: int = None,
        min_row: int = None,
        max_row: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """
        Produces cells from a table, by column. Specify the iteration range using
        the indexes of the rows and columns.

        Args:
            min_col: smallest row index (zero indexed)
            max_col: largest row (zero indexed)
            min_row: smallest row index (zero indexed)
            max_row: largest row (zero indexed)
            values_only: return cell values rather than Cell objects

        Returns:
            generator: tuple of cells

        Raises:
            IndexError: row or column values are out of range for the table
        """
        min_row = min_row or 0
        max_row = max_row or self.num_rows - 1
        min_col = min_col or 0
        max_col = max_col or self.num_cols - 1

        if min_row < 0:
            raise IndexError(f"row {min_row} out of range")
        if max_row > self.num_rows:
            raise IndexError(f"row {max_row} out of range")
        if min_col < 0:
            raise IndexError(f"column {min_col} out of range")
        if max_col > self.num_cols:
            raise IndexError(f"column {max_col} out of range")

        rows = self.rows()
        for col_num in range(min_col, max_col + 1):
            if values_only:
                yield tuple(row[col_num].value for row in rows[min_row : max_row + 1])
            else:
                yield tuple(row[col_num] for row in rows[min_row : max_row + 1])
