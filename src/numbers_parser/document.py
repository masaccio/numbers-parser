from typing import Generator, Tuple, Union
from warnings import warn

from numbers_parser.cell import Border, Cell, MergedCell, Style, xl_cell_to_rowcol
from numbers_parser.cell_storage import CellStorage
from numbers_parser.constants import (
    DEFAULT_COLUMN_COUNT,
    DEFAULT_NUM_HEADERS,
    DEFAULT_ROW_COUNT,
    MAX_COL_COUNT,
    MAX_HEADER_COUNT,
    MAX_ROW_COUNT,
)
from numbers_parser.containers import ItemsList
from numbers_parser.file import write_numbers_file
from numbers_parser.model import _NumbersModel
from numbers_parser.numbers_cache import Cacheable, cache


class Document:
    def __init__(  # noqa: PLR0913
        self,
        filename: str = None,
        sheet_name: str = None,
        table_name: str = None,
        num_header_rows: int = None,
        num_header_cols: int = None,
        num_rows: int = None,
        num_cols: int = None,
    ):
        if filename is not None and (
            (sheet_name is not None)
            or (table_name is not None)
            or (num_header_rows is not None)
            or (num_header_cols is not None)
            or (num_rows is not None)
            or (num_cols is not None)
        ):
            warn(
                "can't set table/sheet attributes on load of existing document",
                RuntimeWarning,
                stacklevel=2,
            )

        self._model = _NumbersModel(filename)
        refs = self._model.sheet_ids()
        self._sheets = ItemsList(self._model, refs, Sheet)

        if filename is None:
            if sheet_name is not None:
                self.sheets[0].name = sheet_name
            table = self.sheets[0].tables[0]
            if table_name is not None:
                table.name = table_name

            if num_header_rows is None:
                num_header_rows = DEFAULT_NUM_HEADERS
            if num_header_cols is None:
                num_header_cols = DEFAULT_NUM_HEADERS
            if num_rows is None:
                num_rows = DEFAULT_ROW_COUNT
            if num_cols is None:
                num_cols = DEFAULT_COLUMN_COUNT

            # Table starts as 1x1 with no headers
            table.add_row(num_rows - 1)
            table.num_header_rows = num_header_rows
            table.add_column(num_cols - 1)
            table.num_header_cols = num_header_cols

    @property
    def sheets(self) -> list:
        """Return a list of all sheets in the document."""
        return self._sheets

    @property
    def styles(self) -> list:
        """Return a list of styles available in the document."""
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
        the next available numbered sheet will be generated.
        """
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
        provided, the next available numbered style will be generated.
        """
        if "name" in kwargs and kwargs["name"] is not None and kwargs["name"] in self._model.styles:
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
        """Return the sheets name."""
        return self._model.sheet_name(self._sheet_id)

    @name.setter
    def name(self, value):
        """Set the sheet's name."""
        self._model.sheet_name(self._sheet_id, value)

    def add_table(  # noqa: PLR0913
        self,
        table_name=None,
        x=None,
        y=None,
        num_rows=DEFAULT_ROW_COUNT,
        num_cols=DEFAULT_COLUMN_COUNT,
    ) -> object:
        """Add a new table to the current sheet. If no table name is provided,
        the next available numbered table will be generated.
        """
        from_table_id = self._tables[-1]._table_id
        return self._add_table(table_name, from_table_id, x, y, num_rows, num_cols)

    def _add_table(  # noqa: PLR0913
        self, table_name, from_table_id, x, y, num_rows, num_cols
    ) -> object:  # noqa: PLR0913
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


class Table(Cacheable):
    def __init__(self, model, table_id):
        super().__init__()
        self._model = model
        self._table_id = table_id
        self.num_rows = self._model.number_of_rows(self._table_id)
        self.num_cols = self._model.number_of_columns(self._table_id)
        # Cache all data now to facilite write(). Performance impact
        # of computing all cells is minimal compared to file IO
        self._data = []
        self._model.set_table_data(table_id, self._data)
        merge_cells = self._model.merge_cells(table_id)

        for row_num in range(self.num_rows):
            self._data.append([])
            for col_num in range(self.num_cols):
                cell_storage = model.table_cell_decode(table_id, row_num, col_num)
                if cell_storage is None:
                    if merge_cells.is_merge_reference((row_num, col_num)):
                        cell = Cell.merged_cell(table_id, row_num, col_num, model)
                    else:
                        cell = Cell.empty_cell(table_id, row_num, col_num, model)
                else:
                    cell = Cell.from_storage(cell_storage)
                self._data[row_num].append(cell)

    @property
    def name(self) -> str:
        """Return the table's name."""
        return self._model.table_name(self._table_id)

    @name.setter
    def name(self, value: str):
        """Set the table's name."""
        self._model.table_name(self._table_id, value)

    @property
    def table_name_enabled(self):
        return self._model.table_name_enabled(self._table_id)

    @table_name_enabled.setter
    def table_name_enabled(self, enabled):
        self._model.table_name_enabled(self._table_id, enabled)

    @property
    def num_header_rows(self) -> int:
        """Return the number of header rows."""
        return self._model.num_header_rows(self._table_id)

    @num_header_rows.setter
    def num_header_rows(self, num_headers: int):
        """Return the number of header rows."""
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_rows:
            raise ValueError("Number of headers cannot exceed the number of rows")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} rows")
        return self._model.num_header_rows(self._table_id, num_headers)

    @property
    def num_header_cols(self) -> int:
        """Return the number of header columns."""
        return self._model.num_header_cols(self._table_id)

    @num_header_cols.setter
    def num_header_cols(self, num_headers: int):
        """Return the number of header columns."""
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_cols:
            raise ValueError("Number of headers cannot exceed the number of columns")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} columns")
        return self._model.num_header_cols(self._table_id, num_headers)

    @property
    def height(self) -> int:
        """Return the table's height in points."""
        return self._model.table_height(self._table_id)

    @property
    def width(self) -> int:
        """Return the table's width in points."""
        return self._model.table_width(self._table_id)

    def row_height(self, row_num: int, height: int = None) -> int:
        """Return the height of a table row. Set the height if not None."""
        return self._model.row_height(self._table_id, row_num, height)

    def col_width(self, col_num: int, width: int = None) -> int:
        """Return the width of a table column. Set the width if not None."""
        return self._model.col_width(self._table_id, col_num, width)

    @property
    def coordinates(self) -> Tuple[float]:
        """Return the table's x,y offsets in points."""
        return self._model.table_coordinates(self._table_id)

    def rows(self, values_only: bool = False) -> list:
        """Return all rows of cells for the Table.

        Args:
            values_only: if True, return cell values instead of Cell objects

        Returns:
            rows: list of rows; each row is a list of Cell objects
        """
        if values_only:
            return [[cell.value for cell in row] for row in self._data]
        else:
            return self._data

    @property
    @cache(num_args=0)
    def merge_ranges(self) -> list:
        merge_cells = self._model.merge_cells(self._table_id).merge_cell_names()
        return sorted(set(list(merge_cells)))

    def cell(self, *args) -> Union[Cell, MergedCell]:
        if isinstance(args[0], str):
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

    def iter_rows(  # noqa: PLR0913
        self,
        min_row: int = None,
        max_row: int = None,
        min_col: int = None,
        max_col: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """Produces cells from a table, by row. Specify the iteration range using
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

    def iter_cols(  # noqa: PLR0913
        self,
        min_col: int = None,
        max_col: int = None,
        min_row: int = None,
        max_row: int = None,
        values_only: bool = False,
    ) -> Generator[tuple, None, None]:
        """Produces cells from a table, by column. Specify the iteration range using
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
        the table with empty cells if outside current bounds.
        """
        if isinstance(args[0], str):
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

        for _ in range(self.num_rows, row_num + 1):
            self.add_row()

        for _ in range(self.num_cols, col_num + 1):
            self.add_column()

        return (row_num, col_num) + tuple(values)

    def write(self, *args, style=None, formatting=None):
        # TODO: write needs to retain/init the border
        (row_num, col_num, value) = self._validate_cell_coords(*args)
        self._data[row_num][col_num] = Cell.from_value(row_num, col_num, value)
        self._data[row_num][col_num]._storage = CellStorage(
            self._model, self._table_id, None, row_num, col_num
        )
        merge_cells = self._model.merge_cells(self._table_id)
        self._data[row_num][col_num]._table_id = self._table_id
        self._data[row_num][col_num]._model = self._model
        self._data[row_num][col_num]._set_merge(merge_cells.get((row_num, col_num)))

        if style is not None:
            self.set_cell_style(row_num, col_num, style)
        if formatting is not None:
            self.set_cell_formatting(row_num, col_num, formatting)

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

    def set_cell_formatting(self, *args):
        (row_num, col_num, formatting) = self._validate_cell_coords(*args)
        if not isinstance(formatting, dict):
            raise TypeError("formatting values must be a dict")

        self._data[row_num][col_num].set_formatting(formatting)

    def add_row(self, num_rows=1):
        row = [
            Cell.empty_cell(self._table_id, self.num_rows - 1, col_num, self._model)
            for col_num in range(self.num_cols)
        ]
        for _ in range(num_rows):
            self._data.append(row.copy())
            self.num_rows += 1
            self._model.number_of_rows(self._table_id, self.num_rows)

    def add_column(self, num_cols=1):
        for _ in range(num_cols):
            for row_num in range(self.num_rows):
                self._data[row_num].append(
                    Cell.empty_cell(self._table_id, row_num, self.num_cols - 1, self._model)
                )
            self.num_cols += 1
            self._model.number_of_columns(self._table_id, self.num_cols)

    def merge_cells(self, cell_range):
        """Convert a cell range or list of cell ranges into merged cells."""
        if isinstance(cell_range, list):
            for x in cell_range:
                self.merge_cells(x)
        else:
            (start_cell_ref, end_cell_ref) = cell_range.split(":")
            (row_start, col_start) = xl_cell_to_rowcol(start_cell_ref)
            (row_end, col_end) = xl_cell_to_rowcol(end_cell_ref)
            num_rows = row_end - row_start + 1
            num_cols = col_end - col_start + 1

            merge_cells = self._model.merge_cells(self._table_id)
            merge_cells.add_anchor(row_start, col_start, (num_rows, num_cols))
            for row_num in range(row_start + 1, row_end + 1):
                for col_num in range(col_start + 1, col_end + 1):
                    self._data[row_num][col_num] = MergedCell(row_num, col_num)
                    merge_cells.add_reference(
                        row_num, col_num, (row_start, col_start, row_end, col_end)
                    )

            for row_num, row in enumerate(self._data):
                for col_num, cell in enumerate(row):
                    cell._set_merge(merge_cells.get((row_num, col_num)))

    def set_cell_border(self, *args):
        (row_num, col_num, *args) = self._validate_cell_coords(*args)
        if len(args) == 2:
            (side, border_value) = args
            length = 1
        elif len(args) == 3:
            (side, border_value, length) = args
        else:
            raise TypeError("invalid number of arguments to border_value()")

        if not isinstance(border_value, Border):
            raise TypeError("border value must be a Border object")

        if not isinstance(length, int):
            raise TypeError("border length must be an int")

        if isinstance(side, list):
            for s in side:
                self.set_cell_border(row_num, col_num, s, border_value, length)
            return

        if self._data[row_num][col_num].is_merged and side in ["bottom", "right"]:
            warn(
                f"cell [{row_num},{col_num}] is merged; {side} border not set",
                RuntimeWarning,
                stacklevel=2,
            )
            return

        self._model.extract_strokes(self._table_id)

        if side in ["top", "bottom"]:
            for border_col_num in range(col_num, col_num + length):
                self._model.set_cell_border(
                    self._table_id, row_num, border_col_num, side, border_value
                )
        elif side in ["left", "right"]:
            for border_row_num in range(row_num, row_num + length):
                self._model.set_cell_border(
                    self._table_id, border_row_num, col_num, side, border_value
                )
        else:
            raise TypeError("side must be a valid border segment")

        self._model.add_stroke(self._table_id, row_num, col_num, side, border_value, length)
