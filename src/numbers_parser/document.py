from datetime import timedelta, datetime
from functools import lru_cache
from typing import Union, Generator

from numbers_parser.containers import ItemsList
from numbers_parser.model import _NumbersModel
from numbers_parser.file import write_numbers_file
from numbers_parser.cell import (
    EmptyCell,
    BoolCell,
    NumberCell,
    TextCell,
    DurationCell,
    DateCell,
)
from numbers_parser.cell import (
    Cell,
    MergedCell,
    xl_cell_to_rowcol,
    xl_range,
)

_EXPERIMENTAL_NUMBERS_PARSER = False


class Document:
    def __init__(self, filename=None):
        self._model = _NumbersModel(filename)
        refs = self._model.sheet_ids()
        self._sheets = ItemsList(self._model, refs, Sheet)

    @property
    def sheets(self):
        return self._sheets

    def save(self, filename):
        for sheet in self.sheets:
            for table in sheet.tables:
                self._model.recalculate_table_data(table._table_id, table._data)
        write_numbers_file(filename, self._model.file_store)

    def add_sheet(self, sheet_name=None, table_name=None) -> object:
        """Add a new sheet to the current document. If no sheet name is provided,
        the next available numbered sheet will be generated"""
        if not _EXPERIMENTAL_NUMBERS_PARSER:
            raise AttributeError("'Document' object has no attribute 'add_sheet'")

        if sheet_name is not None:
            if sheet_name in self._sheets:
                raise IndexError(f"sheet '{sheet_name}' already exists")
        else:
            sheet_num = 1
            while f"sheet {sheet_num}" in self._sheets:
                sheet_num += 1
            sheet_name = f"Sheet {sheet_num}"

        if table_name is None:
            table_name = "Table 1"

        new_sheet_id = self._model.add_sheet(sheet_name, table_name, self._sheets[-1])
        self._sheets.append(Sheet(self._model, new_sheet_id))
        return self._sheets[-1]


class Sheet:
    def __init__(self, model, sheet_id):
        self._sheet_id = sheet_id
        self._model = model
        refs = self._model.table_ids(self._sheet_id)
        self._tables = ItemsList(self._model, refs, Table)

    @property
    def tables(self):
        return self._tables

    @property
    def name(self):
        return self._model.sheet_name(self._sheet_id)

    @name.setter
    def name(self, value):
        self._model.sheet_name(self._sheet_id, value)

    def add_table(self, table_name=None) -> object:
        """Add a new table to the current sheet. If no sheet name is provided,
        the next available numbered sheet will be generated"""
        if not _EXPERIMENTAL_NUMBERS_PARSER:
            raise AttributeError("'Sheet' object has no attribute 'add_table'")
        if table_name is not None:
            if table_name in self._tables:
                raise IndexError(f"table '{table_name}' already exists")
        else:
            table_num = 1
            while f"table {table_num}" in self._tables:
                table_num += 1
            table_name = f"Table {table_num}"

        new_table_id = self._model.add_table(self._sheet_id, table_name)
        self._tables.append(Table(self._model, new_table_id))
        return self._tables[-1]


class Table:
    def __init__(self, model, table_id):
        super(Table, self).__init__()
        self._model = model
        self._table_id = table_id
        self.num_rows = self._model.number_of_rows(self._table_id)
        self.num_cols = self._model.number_of_columns(self._table_id)
        # Cache all data now to facilite write(). Performance impact
        # of computing all cells is minimal compared to file IO
        self._data = [
            [
                Cell.factory(self._model, self._table_id, row_num, col_num)
                for col_num in range(self.num_cols)
            ]
            for row_num in range(self.num_rows)
        ]

    @property
    def name(self):
        return self._model.table_name(self._table_id)

    @name.setter
    def name(self, value):
        self._model.table_name(self._table_id, value)

    @lru_cache(maxsize=None)
    def rows(self, values_only: bool = False) -> list:
        """
        Return all rows of cells for the Table.

        Args:
            values_only: if True, return cell values instead of Cell objects

        Returns:
            rows: list of rows; each row is a list of Cell objects
        """
        if values_only:
            rows = [[cell.value for cell in row] for row in self._data]
            return rows
        else:
            return self._data

    @property
    @lru_cache(maxsize=None)
    def merge_ranges(self) -> list:
        merge_cells = self._model.merge_cell_ranges(self._table_id)
        ranges = [xl_range(*r["rect"]) for r in merge_cells.values()]
        return sorted(set(list(ranges)))

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

        return self._data[row_num][col_num]

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
                yield tuple(cell.value for cell in rows[row_num][min_col : max_col + 1])
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

    def write(self, *args):
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
            value = args[1]
        elif len(args) < 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args[0:2]
            value = args[2]

        for row in range(self.num_rows, row_num + 1):
            self.add_row()

        for row in range(self.num_cols, col_num + 1):
            self.add_column()

        if type(value) == str:
            self._data[row_num][col_num] = TextCell(row_num, col_num, value)
        elif type(value) == int or type(value) == float:
            self._data[row_num][col_num] = NumberCell(row_num, col_num, value)
        elif type(value) == bool:
            self._data[row_num][col_num] = BoolCell(row_num, col_num, value)
        elif type(value) == datetime:
            self._data[row_num][col_num] = DateCell(row_num, col_num, value)
        elif type(value) == timedelta:
            self._data[row_num][col_num] = DurationCell(row_num, col_num, value)
        else:
            raise ValueError(
                "Can't determine cell type from type " + type(value).__name__
            )

    def add_row(self, num_rows=1):
        row = [
            EmptyCell(self.num_rows - 1, col_num, None)
            for col_num in range(self.num_cols)
        ]
        for _ in range(num_rows):
            self._data.append(row.copy())
            self.num_rows += 1
            self._model.number_of_rows(self._table_id, self.num_rows)

    def add_column(self, num_cols=1):
        for _ in range(num_cols):
            for row_num in range(self.num_rows):
                self._data[row_num].append(EmptyCell(row_num, self.num_cols - 1, None))
            self.num_cols += 1
            self._model.number_of_columns(self._table_id, self.num_cols)

    def merge_cells(self, cell_range):
        if isinstance(cell_range, list):
            for x in cell_range:
                self.merge_cells(x)
        else:
            (start_cell_ref, end_cell_ref) = cell_range.split(":")
            (row_start, col_start) = xl_cell_to_rowcol(start_cell_ref)
            (row_end, col_end) = xl_cell_to_rowcol(end_cell_ref)
            num_rows = row_end - row_start + 1
            num_cols = col_end - col_start + 1

            merge_ranges = self._model._merge_cells[self._table_id]
            merge_ranges[(row_start, col_start)] = {
                "merge_type": "source",
                "size": (num_rows, num_cols),
            }
            for row_num in range(row_start + 1, row_end + 1):
                for col_num in range(col_start + 1, col_end + 1):
                    merge_ranges[(row_num, col_num)] = {
                        "merge_type": "ref",
                        "rect": (row_start, col_start, row_end, col_end),
                        "size": (num_rows, num_cols),
                    }
