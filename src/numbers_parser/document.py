from datetime import datetime as builtin_datetime, timedelta as builtin_timedelta
from functools import lru_cache
from pendulum import DateTime, Duration, instance as pendulum_instance
from typing import Union, Generator, Tuple

from numbers_parser.containers import ItemsList
from numbers_parser.model import _NumbersModel
from numbers_parser.file import write_numbers_file
from numbers_parser.constants import MAX_ROW_COUNT, MAX_COL_COUNT, MAX_HEADER_COUNT
from numbers_parser.cell import (
    EmptyCell,
    BoolCell,
    NumberCell,
    TextCell,
    DurationCell,
    DateCell,
    Cell,
    MergedCell,
    xl_cell_to_rowcol,
    xl_range,
    Style,
    Border,
)
from numbers_parser.cell_storage import CellStorage
from numbers_parser.constants import DEFAULT_COLUMN_COUNT, DEFAULT_ROW_COUNT


class Document:
    def __init__(self, filename=None):
        self._model = _NumbersModel(filename)
        refs = self._model.sheet_ids()
        self._sheets = ItemsList(self._model, refs, Sheet)

    @property
    def sheets(self) -> list:
        """Return a list of all sheets in the document"""
        return self._sheets

    @property
    def styles(self) -> list:
        """Return a list of styles available in the document"""
        return self._model.styles

    def save(self, filename):
        for sheet in self.sheets:
            for table in sheet.tables:
                self._model.recalculate_table_data(table._table_id, table._data)
        write_numbers_file(filename, self._model.file_store)

    def add_sheet(
        self,
        sheet_name=None,
        table_name=None,
        num_rows=DEFAULT_ROW_COUNT,
        num_cols=DEFAULT_COLUMN_COUNT,
    ) -> object:
        """Add a new sheet to the current document. If no sheet name is provided,
        the next available numbered sheet will be generated"""
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

        prev_table_id = self._sheets[-1]._tables[0]._table_id
        new_sheet_id = self._model.add_sheet(sheet_name)
        new_sheet = Sheet(self._model, new_sheet_id)
        new_sheet._tables.append(
            Table(
                self._model,
                self._model.add_table(
                    new_sheet_id, table_name, prev_table_id, 0, 0, num_rows, num_cols
                ),
            )
        )
        self._sheets.append(new_sheet)

        return new_sheet

    def add_style(self, **kwargs) -> Style:
        """Add a new style to the current document. If no style name is
        provided, the next available numbered style will be generated"""
        if "name" in kwargs and kwargs["name"] is not None:
            if kwargs["name"] in self._model.styles.keys():
                raise IndexError(f"style '{kwargs['name']}' already exists")
        style = Style(**kwargs)
        if style.name is None:
            style.name = self._model.custom_style_name()
        style._update_styles = True
        self._model.styles[style.name] = style
        return style


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
        """Return the sheets name"""
        return self._model.sheet_name(self._sheet_id)

    @name.setter
    def name(self, value):
        """Set the sheet's name"""
        self._model.sheet_name(self._sheet_id, value)

    def add_table(
        self,
        table_name=None,
        x=None,
        y=None,
        num_rows=DEFAULT_ROW_COUNT,
        num_cols=DEFAULT_COLUMN_COUNT,
    ) -> object:
        """Add a new table to the current sheet. If no table name is provided,
        the next available numbered table will be generated"""

        from_table_id = self._tables[-1]._table_id
        return self._add_table(table_name, from_table_id, x, y, num_rows, num_cols)

    def _add_table(self, table_name, from_table_id, x, y, num_rows, num_cols) -> object:
        if table_name is not None:
            if table_name in self._tables:
                raise IndexError(f"table '{table_name}' already exists")
        else:
            table_num = 1
            while f"table {table_num}" in self._tables:
                table_num += 1
            table_name = f"Table {table_num}"

        new_table_id = self._model.add_table(
            self._sheet_id, table_name, from_table_id, x, y, num_rows, num_cols
        )
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
        self._data = []
        self._model.set_table_data(table_id, self._data)
        for row_num in range(self.num_rows):
            self._data.append([])
            for col_num in range(self.num_cols):
                cell_storage = model.table_cell_decode(table_id, row_num, col_num)
                if cell_storage is None:
                    cell = Cell.empty_cell(table_id, row_num, col_num, model)
                else:
                    cell = Cell.from_storage(cell_storage)
                self._data[row_num].append(cell)

    @property
    def name(self) -> str:
        """Return the table's name"""
        return self._model.table_name(self._table_id)

    @name.setter
    def name(self, value: str):
        """Set the table's name"""
        self._model.table_name(self._table_id, value)

    @property
    def num_header_rows(self) -> int:
        """Return the number of header rows"""
        return self._model.num_header_rows(self._table_id)

    @num_header_rows.setter
    def num_header_rows(self, num_headers: int):
        """Return the number of header rows"""
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_cols:
            raise ValueError("Number of headers cannot exceed the number of rows")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} rows")
        return self._model.num_header_rows(self._table_id, num_headers)

    @property
    def num_header_cols(self) -> int:
        """Return the number of header columns"""
        return self._model.num_header_cols(self._table_id)

    @num_header_cols.setter
    def num_header_cols(self, num_headers: int):
        """Return the number of header columns"""
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_cols:
            raise ValueError("Number of headers cannot exceed the number of columns")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} columns")
        return self._model.num_header_cols(self._table_id, num_headers)

    @property
    def height(self) -> int:
        """Return the table's height in points"""
        return self._model.table_height(self._table_id)

    @property
    def width(self) -> int:
        """Return the table's width in points"""
        return self._model.table_width(self._table_id)

    def row_height(self, row_num: int, height: int = None) -> int:
        """Return the height of a table row. Set the height if not None"""
        return self._model.row_height(self._table_id, row_num, height)

    def col_width(self, col_num: int, width: int = None) -> int:
        """Return the width of a table column. Set the width if not None"""
        return self._model.col_width(self._table_id, col_num, width)

    @property
    def coordinates(self) -> Tuple[float]:
        """Return the table's x,y offsets in points"""
        return self._model.table_coordinates(self._table_id)

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
            raise IndexError(f"column {col_num} out of range")

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

    def _validate_cell_coords(self, *args):
        """Check first arguments are value cell references and pad
        the table with empty cells if outside current bounds"""
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
            values = args[1:]
        elif len(args) < 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args[0:2]
            values = args[2:]

        if row_num >= MAX_ROW_COUNT:
            raise IndexError(f"{row_num} exceeds maximum row {MAX_ROW_COUNT-1}")
        if col_num >= MAX_COL_COUNT:
            raise IndexError(f"{col_num} exceeds maximum column {MAX_COL_COUNT-1}")

        for row in range(self.num_rows, row_num + 1):
            self.add_row()

        for row in range(self.num_cols, col_num + 1):
            self.add_column()

        return (row_num, col_num) + tuple(values)

    def write(self, *args, style=None):
        (row_num, col_num, value) = self._validate_cell_coords(*args)

        if type(value) == str:
            self._data[row_num][col_num] = TextCell(row_num, col_num, value)
        elif type(value) == int or type(value) == float:
            self._data[row_num][col_num] = NumberCell(row_num, col_num, value)
        elif type(value) == bool:
            self._data[row_num][col_num] = BoolCell(row_num, col_num, value)
        elif type(value) == builtin_datetime or type(value) == DateTime:
            self._data[row_num][col_num] = DateCell(row_num, col_num, pendulum_instance(value))
        elif type(value) == builtin_timedelta or type(value) == Duration:
            self._data[row_num][col_num] = DurationCell(row_num, col_num, value)
        else:
            raise ValueError("Can't determine cell type from type " + type(value).__name__)
        self._data[row_num][col_num]._storage = CellStorage(
            self._model, self._table_id, None, row_num, col_num
        )
        self._data[row_num][col_num]._table_id = self._table_id
        self._data[row_num][col_num]._model = self._model

        if style is not None:
            self.set_cell_style(row_num, col_num, style)

    def set_cell_style(self, *args):
        (row_num, col_num, style) = self._validate_cell_coords(*args)
        if isinstance(style, Style):
            self._data[row_num][col_num]._style = style
        elif isinstance(style, str):
            if style not in self._model.styles:
                raise IndexError(f"style '{style}' does not exist")
            self._data[row_num][col_num]._style = self._model.styles[style]
        else:
            raise TypeError("style must be a Style object or style name")

    def add_row(self, num_rows=1):
        row = [EmptyCell(self.num_rows - 1, col_num) for col_num in range(self.num_cols)]
        for _ in range(num_rows):
            self._data.append(row.copy())
            self.num_rows += 1
            self._model.number_of_rows(self._table_id, self.num_rows)

    def add_column(self, num_cols=1):
        for _ in range(num_cols):
            for row_num in range(self.num_rows):
                self._data[row_num].append(EmptyCell(row_num, self.num_cols - 1))
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

    def add_border(self, *args):
        (row_num, col_num, *args) = self._validate_cell_coords(*args)
        if len(args) == 2:
            (side, border_value) = args
            length = 0
        elif len(args) == 3:
            (side, border_value, length) = args
        else:
            raise TypeError("invalid number of arguments to border_value()")

        if not isinstance(border_value, Border):
            raise TypeError("border value must be a Border object")

        if not isinstance(length, int):
            raise TypeError("border length must be an int")

        if side == "top" or side == "bottom":
            for border_col_num in range(col_num, col_num + length):
                self._model.set_cell_border(
                    self._table_id, row_num, border_col_num, side, border_value
                )
        elif side == "left" or side == "right":
            for border_row_num in range(row_num, row_num + length):
                self._model.set_cell_border(
                    self._table_id, border_row_num, col_num, side, border_value
                )
        else:
            raise TypeError("side must be a valid border segment")

        self._model.update_strokes(self._table_id, row_num, col_num, side, length)
