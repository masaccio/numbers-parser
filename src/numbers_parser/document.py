from typing import Dict, Iterator, List, Optional, Tuple, Union
from warnings import warn

from pendulum import DateTime, Duration

from numbers_parser.cell import (
    BackgroundImage,
    Border,
    Cell,
    ControlFormattingType,
    CustomFormatting,
    CustomFormattingType,
    Formatting,
    FormattingType,
    MergedCell,
    Style,
    TextCell,
    UnsupportedWarning,
    xl_cell_to_rowcol,
    xl_range,
)
from numbers_parser.cell_storage import CellStorage
from numbers_parser.constants import (
    CUSTOM_FORMATTING_ALLOWED_CELLS,
    DEFAULT_COLUMN_COUNT,
    DEFAULT_ROW_COUNT,
    FORMATTING_ACTION_CELLS,
    FORMATTING_ALLOWED_CELLS,
    MAX_COL_COUNT,
    MAX_HEADER_COUNT,
    MAX_ROW_COUNT,
)
from numbers_parser.containers import ItemsList
from numbers_parser.file import write_numbers_file
from numbers_parser.model import _NumbersModel
from numbers_parser.numbers_cache import Cacheable

__all__ = ["Document", "Sheet", "Table"]


class Sheet:
    pass


class Table:
    pass


class Document:
    """
    Create an instance of a new Numbers document.

    If ``filename`` is ``None``, an empty document is created using the defaults
    defined by the class constructor. You can optionionally override these
    defaults at object construction time.

    Parameters
    ----------
    filename: str, optional
        Apple Numbers document to read.
    sheet_name: *str*, *optional*, *default*: ``Sheet 1``
        Name of the first sheet in a new document
    table_name: *str*, *optional*, *default*: ``Table 1``
        Name of the first table in the first sheet of a new
    num_header_rows: int, optional, default: 1
        Number of header rows in the first table of a new document.
    num_header_cols: int, optional, default: 1
        Number of header columns in the first table of a new document.
    num_rows: int, optional, default: 12
        Number of rows in the first table of a new document.
    num_cols: int, optional, default: 8
        Number of columns in the first table of a new document.

    Raises
    ------
    IndexError:
        If the sheet name already exists in the document.
    IndexError:
        If the table name already exists in the first sheet.
    """

    def __init__(  # noqa: PLR0913
        self,
        filename: Optional[str] = None,
        sheet_name: Optional[str] = "Sheet 1",
        table_name: Optional[str] = "Table 1",
        num_header_rows: Optional[int] = 1,
        num_header_cols: Optional[int] = 1,
        num_rows: Optional[int] = DEFAULT_ROW_COUNT,
        num_cols: Optional[int] = DEFAULT_COLUMN_COUNT,
    ):
        self._model = _NumbersModel(filename)
        refs = self._model.sheet_ids()
        self._sheets = ItemsList(self._model, refs, Sheet)

        if filename is None:
            self.sheets[0].name = sheet_name
            table = self.sheets[0].tables[0]
            table.name = table_name

            # Table starts as 1x1 with no headers
            table.add_row(num_rows - 1)
            table.num_header_rows = num_header_rows
            table.add_column(num_cols - 1)
            table.num_header_cols = num_header_cols

    @property
    def sheets(self) -> List[Sheet]:
        """List[:class:`Sheet`]: A list of sheets in the document."""
        return self._sheets

    @property
    def styles(self) -> Dict[str, Style]:
        """Dict[str, :class:`Style`]: A dict mapping style names to to the corresponding style."""
        return self._model.styles

    @property
    def custom_formats(self) -> Dict[str, CustomFormatting]:
        """
        Dict[str, :class:`CustomFormatting`]: A dict mapping custom format names
        to the corresponding custom format.
        """
        return self._model.custom_formats

    def save(self, filename: str) -> None:
        """
        Save the document in the specified filename.

        Parameters
        ----------
        filename: str
            The path to save the document to. If the file already exists,
            it will be overwritten.
        """
        for sheet in self.sheets:
            for table in sheet.tables:
                if self._model.is_a_pivot_table(table._table_id):
                    table_name = self._model.table_name(table._table_id)
                    warn(
                        f"Not modifying pivot table '{table_name}'",
                        UnsupportedWarning,
                        stacklevel=2,
                    )
                else:
                    self._model.recalculate_table_data(table._table_id, table._data)
        write_numbers_file(filename, self._model.file_store)

    def add_sheet(
        self,
        sheet_name: Optional[str] = None,
        table_name: Optional[str] = "Table 1",
        num_rows: Optional[int] = DEFAULT_ROW_COUNT,
        num_cols: Optional[int] = DEFAULT_COLUMN_COUNT,
    ) -> None:
        """
        Add a new sheet to the current document.

        If no sheet name is provided, the next available numbered sheet
        will be generated in the series ``Sheet 1``, ``Sheet 2``, etc.

        Parameters
        ----------
        sheet_name: str, optional
            The name of the sheet to add to the document
        table_name: *str*, *optional*, *default*: ``Table 1``
            The name of the table created in the new sheet
        num_rows: int, optional, default: 12
            The number of columns in the newly created table
        num_cols: int, optional, default: 8
            The number of columns in the newly created table

        Raises:
            IndexError: If the sheet name already exists in the document.
        """
        if sheet_name is not None:
            if sheet_name in self._sheets:
                raise IndexError(f"sheet '{sheet_name}' already exists")
        else:
            sheet_num = 1
            while f"sheet {sheet_num}" in self._sheets:
                sheet_num += 1
            sheet_name = f"Sheet {sheet_num}"

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

    def add_style(self, **kwargs) -> Style:
        r"""
        Add a new style to the current document.

        If no style name is provided, the next available numbered style
        will be generated in the series ``Custom Style 1``, ``Custom Style 2``, etc.

        .. code-block:: python

            red_text = doc.add_style(
                name="Red Text",
                font_name="Lucida Grande",
                font_color=RGB(230, 25, 25),
                font_size=14.0,
                bold=True,
                italic=True,
                alignment=Alignment("right", "top"),
            )
            table.write("B2", "Red", style=red_text)
            table.set_cell_style("C2", red_text)

        Parameters
        ----------
        kwargs: dict, optional
            Key-value pairs to pass to the :py:class:`~numbers_parser.Style` constructor.

        Raises
        ------
        TypeError:
            If ``font_size`` is not a ``float``, ``font_name`` is not a ``str``,
            ``bg_image`` is not a :py:class:`~numbers_parser.BackgroundImage`,
            or if any of the ``bool`` parameters are invalid.
        """  # noqa: E501
        if "name" in kwargs and kwargs["name"] is not None and kwargs["name"] in self._model.styles:
            raise IndexError(f"style '{kwargs['name']}' already exists")

        if "bg_image" in kwargs and kwargs["bg_image"] is not None:
            if not isinstance(kwargs["bg_image"], BackgroundImage):
                raise TypeError("bg_image must be a BackgroundImage object")
            self._model.store_image((kwargs["bg_image"].data), kwargs["bg_image"].filename)

        style = Style(**kwargs)
        if style.name is None:
            style.name = self._model.custom_style_name()
        style._update_styles = True
        self._model.styles[style.name] = style
        return style

    def add_custom_format(self, **kwargs) -> CustomFormatting:
        r"""
        Add a new custom format to the current document.

        .. code-block:: python

            long_date = doc.add_custom_format(
                name="Long Date",
                type="date",
                date_time_format="EEEE, d MMMM yyyy"
            )
            table.set_cell_formatting("C1", "custom", format=long_date)

        All custom formatting styles share a name and a type, described in the **Common**
        parameters in the following table. Additional key-value pairs configure the format
        depending upon the value of ``kwargs["type"]``.

        :Common Args:
            * **name** (``str``) – The name of the custom format. If no name is provided,
              one is generated using the scheme ``Custom Format``, ``Custom Format 1``, ``Custom Format 2``, etc.
            * **type** (``str``, *optional*, default: ``number``) – The type of format to
              create:

              * ``"datetime"``: A date and time value with custom formatting.
              * ``"number"``: A decimal number.
              * ``"text"``: A simple text string.

        :``"number"``:
            * **integer_format** (``PaddingType``, *optional*, default: ``PaddingType.NONE``) – How
              to pad integers.
            * **decimal_format** (``PaddingType``, *optional*, default: ``PaddingType.NONE``) – How
              to pad decimals.
            * **num_integers** (``int``, *optional*, default: ``0``) – Integer precision
              when integers are padded.
            * **num_decimals** (``int``, *optional*, default: ``0``) – Integer precision
              when decimals are padded.
            * **show_thousands_separator** (``bool``, *optional*, default: ``False``) – ``True``
              if the number should include a thousands seperator.

        :``"datetime"``:
            * **format** (``str``, *optional*, default: ``"d MMM y"``) – A POSIX strftime-like
              formatting string of `Numbers date/time directives <#datetime-formats>`_.

        :``"text"``:
            * **format** (``str``, *optional*, default: ``"%s"``) – Text format.
              The cell value is inserted in place of %s. Only one substitution is allowed by
              Numbers, and multiple %s formatting references raise a TypeError exception
        """  # noqa: E501
        if (
            "name" in kwargs
            and kwargs["name"] is not None
            and kwargs["name"] in self._model.custom_formats
        ):
            raise IndexError(f"format '{kwargs['name']}' already exists")

        if "type" in kwargs:
            format_type = kwargs["type"].upper()
            try:
                kwargs["type"] = CustomFormattingType[format_type]
            except (KeyError, AttributeError):
                raise TypeError(f"unsuported cell format type '{format_type}'") from None

        custom_format = CustomFormatting(**kwargs)
        if custom_format.name is None:
            custom_format.name = self._model.custom_format_name()
        if custom_format.type == CustomFormattingType.NUMBER:
            self._model.add_custom_decimal_format_archive(custom_format)
        elif custom_format.type == CustomFormattingType.DATETIME:
            self._model.add_custom_datetime_format_archive(custom_format)
        else:
            self._model.add_custom_text_format_archive(custom_format)
        return custom_format


class Sheet:
    def __init__(self, model, sheet_id):
        self._sheet_id = sheet_id
        self._model = model
        refs = self._model.table_ids(self._sheet_id)
        self._tables = ItemsList(self._model, refs, Table)

    @property
    def tables(self) -> List[Table]:
        """List[:class:`Table`]: A list of tables in the sheet."""
        return self._tables

    @property
    def name(self) -> str:
        """str: The name of the sheet."""
        return self._model.sheet_name(self._sheet_id)

    @name.setter
    def name(self, value: str):
        self._model.sheet_name(self._sheet_id, value)

    def add_table(  # noqa: PLR0913
        self,
        table_name: Optional[str] = None,
        x: Optional[float] = None,
        y: Optional[float] = None,
        num_rows: Optional[int] = DEFAULT_ROW_COUNT,
        num_cols: Optional[int] = DEFAULT_COLUMN_COUNT,
    ) -> Table:
        """Add a new table to the current sheet.

        If no table name is provided, the next available numbered table
        will be generated in the series ``Table 1``, ``Table 2``, etc.

        By default, new tables are positioned at a fixed offset below the last
        table vertically in a sheet and on the left side of the sheet. Large
        table headers and captions may result in new tables overlapping existing
        ones. The ``add_table`` method takes optional coordinates for
        positioning a table. A table's height and coordinates can also be
        queried to help aligning new tables:

        .. code:: python

            (x, y) = sheet.table[0].coordinates
            y += sheet.table[0].height + 200.0
            new_table = sheet.add_table("Offset Table", x, y)

        Parameters
        ----------
        table_name: str, optional
            The name of the new table.
        x: float, optional
            The x offset for the table in points.
        y: float, optional
            The y offset for the table in points.
        num_rows: int, optional, default: 12
            The number of rows for the new table.
        num_cols: int, optional, default: 10
            The number of columns for the new table.

        Returns
        -------
        Table
            The newly created table.

        Raises
        ------
            IndexError: If the table name already exists.
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


class Table(Cacheable):  # noqa: F811
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

        for row in range(self.num_rows):
            self._data.append([])
            for col in range(self.num_cols):
                cell_storage = model.table_cell_decode(table_id, row, col)
                if cell_storage is None:
                    if merge_cells.is_merge_reference((row, col)):
                        cell = Cell.merged_cell(table_id, row, col, model)
                    else:
                        cell = Cell.empty_cell(table_id, row, col, model)
                else:
                    cell = Cell.from_storage(cell_storage)
                self._data[row].append(cell)

    @property
    def name(self) -> str:
        """str: The table's name."""
        return self._model.table_name(self._table_id)

    @name.setter
    def name(self, value: str):
        self._model.table_name(self._table_id, value)

    @property
    def table_name_enabled(self) -> bool:
        """bool: ``True`` if the table name is visible, ``False`` otherwise."""
        return self._model.table_name_enabled(self._table_id)

    @table_name_enabled.setter
    def table_name_enabled(self, enabled: bool):
        self._model.table_name_enabled(self._table_id, enabled)

    @property
    def num_header_rows(self) -> int:
        """
        int: The number of header rows.

        Example
        -------

        .. code-block:: python

            # Add an extra header row
            table.num_header_rows += 1

        Raises
        ------
        ValueError:
            If the number of headers is negative, exceeds the number of rows in the
            table, or exceeds Numbers maxinum number of headers (``MAX_HEADER_COUNT``).
        """
        return self._model.num_header_rows(self._table_id)

    @num_header_rows.setter
    def num_header_rows(self, num_headers: int):
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_rows:
            raise ValueError("Number of headers cannot exceed the number of rows")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} rows")
        return self._model.num_header_rows(self._table_id, num_headers)

    @property
    def num_header_cols(self) -> int:
        """
        int: The number of header columns.

        Example
        -------

        .. code-block:: python

            # Add an extra header column
            table.num_header_cols += 1

        Raises
        ------
        ValueError:
            If the number of headers is negative, exceeds the number of rows in the
            table, or exceeds Numbers maxinum number of headers (``MAX_HEADER_COUNT``).
        """
        return self._model.num_header_cols(self._table_id)

    @num_header_cols.setter
    def num_header_cols(self, num_headers: int):
        if num_headers < 0:
            raise ValueError("Number of headers cannot be negative")
        elif num_headers > self.num_cols:
            raise ValueError("Number of headers cannot exceed the number of columns")
        elif num_headers > MAX_HEADER_COUNT:
            raise ValueError(f"Number of headers cannot exceed {MAX_HEADER_COUNT} columns")
        return self._model.num_header_cols(self._table_id, num_headers)

    @property
    def height(self) -> int:
        """int: The table's height in points."""
        return self._model.table_height(self._table_id)

    @property
    def width(self) -> int:
        """int: The table's width in points."""
        return self._model.table_width(self._table_id)

    def row_height(self, row: int, height: int = None) -> int:
        """
        The height of a table row in points.

        .. code-block:: python

            # Double the row's height
            _ = table.row_height(4, table.row_height(4) * 2)

        Parameters
        ----------
        row: int
            The row number (zero indexed).
        height: int
            The height of the row in points. If not ``None``, set the row height.

        Returns
        -------
        int:
            The height of the table row.
        """
        return self._model.row_height(self._table_id, row, height)

    def col_width(self, col: int, width: int = None) -> int:
        """The width of a table column in points.

        Parameters
        ----------
        col: int
            The column number (zero indexed).
        width: int
            The width of the column in points. If not ``None``, set the column width.

        Returns
        -------
        int:
            The width of the table column.
        """
        return self._model.col_width(self._table_id, col, width)

    @property
    def coordinates(self) -> Tuple[float]:
        """Tuple[float]: The table's x, y offsets in points."""
        return self._model.table_coordinates(self._table_id)

    def rows(self, values_only: bool = False) -> Union[List[List[Cell]], List[List[str]]]:
        """Return all rows of cells for the Table.

        Parameters
        ----------
        values_only:
            If ``True``, return cell values instead of :class:`Cell` objects

        Returns
        -------
        List[List[Cell]] | List[List[str]]:
            List of rows; each row is a list of :class:`Cell` objects, or string values.
        """
        if values_only:
            return [[cell.value for cell in row] for row in self._data]
        else:
            return self._data

    @property
    def merge_ranges(self) -> List[str]:
        """List[str]: The merge ranges of cells in A1 notation.

        Example
        -------

        .. code-block:: python

            >>> table.merge_ranges
            ['A4:A10']
            >>> table.cell("A4")
            <numbers_parser.cell.TextCell object at 0x1035f4a90>
            >>> table.cell("A5")
            <numbers_parser.cell.MergedCell object at 0x1035f5310>
        """
        merge_cells = set()
        for row, cells in enumerate(self._data):
            for col, cell in enumerate(cells):
                if cell.is_merged:
                    size = cell.size
                    merge_cells.add(xl_range(row, col, row + size[0] - 1, col + size[1] - 1))
        return sorted(list(merge_cells))

    def cell(self, *args) -> Union[Cell, MergedCell]:
        """
        Return a single cell in the table.

        The ``cell()`` method supports two forms of notation to designate the position
        of cells: **Row-column** notation and **A1** notation:

        .. code-block:: python

            (0, 0)      # Row-column notation.
            ("A1")      # The same cell in A1 notation.

        Parameters
        ----------
        param1: int
            The row number (zero indexed).
        param2: int
            The column number (zero indexed).

        Returns
        -------
        Cell | MergedCell:
            A cell with the base class :class:`Cell` or, if merged, a :class:`MergedCell`.

        Example
        -------

        .. code-block:: python

            >>> doc = Document("mydoc.numbers")
            >>> sheets = doc.sheets
            >>> tables = sheets["Sheet 1"].tables
            >>> table = tables["Table 1"]
            >>> table.cell(1,0)
            <numbers_parser.cell.TextCell object at 0x105a80a10>
            >>> table.cell(1,0).value
            'Debit'
            >>> table.cell("B2")
            <numbers_parser.cell.TextCell object at 0x105a80b90>
            >>> table.cell("B2").value
            1234.50
        """  # noqa: E501
        if isinstance(args[0], str):
            (row, col) = xl_cell_to_rowcol(args[0])
        elif len(args) != 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row, col) = args

        if row >= self.num_rows or row < 0:
            raise IndexError(f"row {row} out of range")
        if col >= self.num_cols or col < 0:
            raise IndexError(f"column {col} out of range")

        return self._data[row][col]

    def iter_rows(  # noqa: PLR0913
        self,
        min_row: Optional[int] = None,
        max_row: Optional[int] = None,
        min_col: Optional[int] = None,
        max_col: Optional[int] = None,
        values_only: Optional[bool] = False,
    ) -> Iterator[Union[Tuple[Cell], Tuple[str]]]:
        """Produces cells from a table, by row.

        Specify the iteration range using the indexes of the rows and columns.

        Parameters
        ----------
        min_row: int, optional
            Starting row number (zero indexed), or ``0`` if ``None``.
        max_row: int, optional
            End row number (zero indexed), or all rows if ``None``.
        min_col: int, optional
            Starting column number (zero indexed) or ``0`` if ``None``.
        max_col: int, optional
            End column number (zero indexed), or all columns if ``None``.
        values_only: bool, optional
            If ``True``, yield cell values rather than :class:`Cell` objects

        Yields
        ------
        Tuple[Cell] | Tuple[str]:
            :class:`Cell` objects or string values for the row

        Raises
        ------
        IndexError:
            If row or column values are out of range for the table

        Example
        -------

        .. code:: python

            for row in table.iter_rows(min_row=2, max_row=7, values_only=True):
                sum += row
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
        for row in range(min_row, max_row + 1):
            if values_only:
                yield tuple(cell.value for cell in rows[row][min_col : max_col + 1])
            else:
                yield tuple(rows[row][min_col : max_col + 1])

    def iter_cols(  # noqa: PLR0913
        self,
        min_col: Optional[int] = None,
        max_col: Optional[int] = None,
        min_row: Optional[int] = None,
        max_row: Optional[int] = None,
        values_only: Optional[bool] = False,
    ) -> Iterator[Union[Tuple[Cell], Tuple[str]]]:
        """Produces cells from a table, by column.

        Specify the iteration range using the indexes of the rows and columns.

        Parameters
        ----------
        min_col: int, optional
            Starting column number (zero indexed) or ``0`` if ``None``.
        max_col: int, optional
            End column number (zero indexed), or all columns if ``None``.
        min_row: int, optional
            Starting row number (zero indexed), or ``0`` if ``None``.
        max_row: int, optional
            End row number (zero indexed), or all rows if ``None``.
        values_only: bool, optional
            If ``True``, yield cell values rather than :class:`Cell` objects.

        Yields
        ------
        Tuple[Cell] | Tuple[str]:
            :class:`Cell` objects or string values for the row

        Raises
        ------
        IndexError:
            If row or column values are out of range for the table

        Example
        =======

        .. code:: python

            for col in table.iter_cols(min_row=2, max_row=7):
                sum += col.value
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
        for col in range(min_col, max_col + 1):
            if values_only:
                yield tuple(row[col].value for row in rows[min_row : max_row + 1])
            else:
                yield tuple(row[col] for row in rows[min_row : max_row + 1])

    def _validate_cell_coords(self, *args):
        if isinstance(args[0], str):
            (row, col) = xl_cell_to_rowcol(args[0])
            values = args[1:]
        elif len(args) < 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row, col) = args[0:2]
            values = args[2:]

        if row >= MAX_ROW_COUNT:
            raise IndexError(f"{row} exceeds maximum row {MAX_ROW_COUNT-1}")
        if col >= MAX_COL_COUNT:
            raise IndexError(f"{col} exceeds maximum column {MAX_COL_COUNT-1}")

        for _ in range(self.num_rows, row + 1):
            self.add_row()

        for _ in range(self.num_cols, col + 1):
            self.add_column()

        return (row, col) + tuple(values)

    def write(self, *args, style: Optional[Union[Style, str, None]] = None) -> None:
        """
        Write a value to a cell and update the style/cell type.

        The ``write()`` method supports two forms of notation to designate the position
        of cells: **Row-column** notation and **A1** notation:

        .. code:: python

            doc = Document("write.numbers")
            sheets = doc.sheets
            tables = sheets[0].tables
            table = tables[0]
            table.write(1, 1, "This is new text")
            table.write("B7", datetime(2020, 12, 25))
            doc.save("new-sheet.numbers")

        Parameters
        ----------

        row: int
            The row number (zero indexed)
        col: int
            The column number (zero indexed)
        value: str | int | float | bool | DateTime | Duration
            The value to write to the cell. The generated cell type is automatically
            created based on the type of ``value``.
        style: Style | str | None
            The name of a document custom style or a :py:class:`~numbers_parser.cell.Style` object.

        Warns
        -----
        RuntimeWarning:
            If the default value is a float that is rounded to the maximum number
            of supported digits.

        Raises
        ------
        IndexError:
            If the style name cannot be foiund in the document.
        TypeError:
            If the style parameter is an invalid type.
        ValueError:
            If the cell type cannot be determined from the type of `param3`.
        """
        # TODO: write needs to retain/init the border
        (row, col, value) = self._validate_cell_coords(*args)
        self._data[row][col] = Cell.from_value(row, col, value)
        storage = CellStorage(self._model, self._table_id, None, row, col)
        storage.update_value(value, self._data[row][col])
        self._data[row][col].update_storage(storage)

        merge_cells = self._model.merge_cells(self._table_id)
        self._data[row][col]._table_id = self._table_id
        self._data[row][col]._model = self._model
        self._data[row][col]._set_merge(merge_cells.get((row, col)))

        if style is not None:
            self.set_cell_style(row, col, style)

    def set_cell_style(self, *args):
        (row, col, style) = self._validate_cell_coords(*args)
        if isinstance(style, Style):
            self._data[row][col]._style = style
        elif isinstance(style, str):
            if style not in self._model.styles:
                raise IndexError(f"style '{style}' does not exist")
            self._data[row][col]._style = self._model.styles[style]
        else:
            raise TypeError("style must be a Style object or style name")

    def add_row(
        self,
        num_rows: Optional[int] = 1,
        start_row: Optional[Union[int, None]] = None,
        default: Optional[Union[str, int, float, bool, DateTime, Duration]] = None,
    ) -> None:
        """
        Add or insert rows to the table.

        Parameters
        ----------
        num_rows: int, optional, default: 1
            The number of rows to add to the table.
        start_row: int, optional, default: None
            The start row number (zero indexed), or ``None`` to add a row to
            the end of the table.
        default: str | int | float | bool | DateTime | Duration, optional, default: None
            The default value for cells. Supported values are those supported by
            :py:meth:`numbers_parser.Table.write` which will determine the new
            cell type.

        Warns
        -----
        RuntimeWarning:
            If the default value is a float that is rounded to the maximum number
            of supported digits.

        Raises
        ------
        IndexError:
            If the start_row is out of range for the table.
        ValueError:
            If the default value is unsupported by :py:meth:`numbers_parser.Table.write`.
        """
        if start_row is not None and (start_row < 0 or start_row >= self.num_rows):
            raise IndexError("Row number not in range for table")

        if start_row is None:
            start_row = self.num_rows
        self.num_rows += num_rows
        self._model.number_of_rows(self._table_id, self.num_rows)

        row = [
            Cell.empty_cell(self._table_id, self.num_rows - 1, col, self._model)
            for col in range(self.num_cols)
        ]
        rows = [row.copy() for _ in range(num_rows)]
        self._data[start_row:start_row] = rows

        if default is not None:
            for row in range(start_row, start_row + num_rows):
                for col in range(self.num_cols):
                    self.write(row, col, default)

    def add_column(
        self,
        num_cols: Optional[int] = 1,
        start_col: Optional[Union[int, None]] = None,
        default: Optional[Union[str, int, float, bool, DateTime, Duration]] = None,
    ) -> None:
        """
        Add or insert columns to the table.

        Parameters
        ----------
        num_cols: int, optional, default: 1
            The number of columns to add to the table.
        start_col: int, optional, default: None
            The start column number (zero indexed), or ``None`` to add a column to
            the end of the table.
        default: str | int | float | bool | DateTime | Duration, optional, default: None
            The default value for cells. Supported values are those supported by
            :py:meth:`numbers_parser.Table.write` which will determine the new
            cell type.

        Warns
        -----
        RuntimeWarning:
            If the default value is a float that is rounded to the maximum number
            of supported digits.

        Raises
        ------
        IndexError:
            If the start_col is out of range for the table.
        ValueError:
            If the default value is unsupported by :py:meth:`numbers_parser.Table.write`.
        """
        if start_col is not None and (start_col < 0 or start_col >= self.num_cols):
            raise IndexError("Column number not in range for table")

        if start_col is None:
            start_col = self.num_cols
        self.num_cols += num_cols
        self._model.number_of_columns(self._table_id, self.num_cols)

        for row in range(self.num_rows):
            cols = [
                Cell.empty_cell(self._table_id, row, col, self._model) for col in range(num_cols)
            ]
            self._data[row][start_col:start_col] = cols

            if default is not None:
                for col in range(start_col, start_col + num_cols):
                    self.write(row, col, default)

    def delete_row(
        self,
        num_rows: Optional[int] = 1,
        start_row: Optional[Union[int, None]] = None,
    ) -> None:
        """
        Delete rows from the table.

        Parameters
        ----------
        num_rows: int, optional, default: 1
            The number of rows to add to the table.
        start_row: int, optional, default: None
            The start row number (zero indexed), or ``None`` to delete rows
            from the end of the table.

        Warns
        -----
        RuntimeWarning:
            If the default value is a float that is rounded to the maximum number
            of supported digits.

        Raises
        ------
        IndexError:
            If the start_row is out of range for the table.
        """
        if start_row is not None and (start_row < 0 or start_row >= self.num_rows):
            raise IndexError("Row number not in range for table")

        if start_row is not None:
            del self._data[start_row : start_row + num_rows]
        else:
            del self._data[-num_rows:]

        self.num_rows -= num_rows
        self._model.number_of_rows(self._table_id, self.num_rows)

    def delete_column(
        self,
        num_cols: Optional[int] = 1,
        start_col: Optional[Union[int, None]] = None,
    ) -> None:
        """
        Add or delete columns columns from the table.

        Parameters
        ----------
        num_cols: int, optional, default: 1
            The number of columns to add to the table.
        start_col: int, optional, default: None
            The start column number (zero indexed), or ``None`` to add delete columns
            from the end of the table.

        Raises
        ------
        IndexError:
            If the start_col is out of range for the table.
        """
        if start_col is not None and (start_col < 0 or start_col >= self.num_cols):
            raise IndexError("Column number not in range for table")

        for row in range(self.num_rows):
            if start_col is not None:
                del self._data[row][start_col : start_col + num_cols]
            else:
                del self._data[row][-num_cols:]

        self.num_cols -= num_cols
        self._model.number_of_columns(self._table_id, self.num_cols)

    def merge_cells(self, cell_range: Union[str, List[str]]) -> None:
        """
        Convert a cell range or list of cell ranges into merged cells.

        Parameters
        ----------
        cell_range: str | List[str]
            Cell range(s) to merge in A1 notation

        Example
        --------
        .. code:: python

            >>> table.cell("B2")
            <numbers_parser.cell.TextCell object at 0x102c0d390>
            >>> table.cell("B2").is_merged
            False
            >>> table.merge_cells("B2:C2")
            >>> table.cell("B2").is_merged
            True
        """
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
            for row in range(row_start + 1, row_end + 1):
                for col in range(col_start + 1, col_end + 1):
                    self._data[row][col] = MergedCell(row, col)
                    merge_cells.add_reference(row, col, (row_start, col_start, row_end, col_end))

            for row, cells in enumerate(self._data):
                for col, cell in enumerate(cells):
                    cell._set_merge(merge_cells.get((row, col)))

    def set_cell_border(self, *args):
        """
        Set the borders for a cell.

        Cell references can be row-column offsers or Excel/Numbers-style A1 notation. Borders
        can be applied to multiple sides of a cell by passing a list of sides. The name(s)
        of the side(s) must be one of ``"top"``, ``"right"``, ``"bottom"`` or ``"left"``.

        Numbers supports different border styles for each cell within a merged cell range
        for those cells that are on the outer part of the merge. ``numbers-parser`` will
        ignore attempts to set these invisible cell edges and issue a ``RuntimeWarning``.

        .. code-block:: python

            # Dashed line for B7's right border
            table.set_cell_border(6, 1, "right", Border(5.0, RGB(29, 177, 0), "dashes"))
            # Solid line starting at B7's left border and running for 3 rows
            table.set_cell_border("B7", "left", Border(8.0, RGB(29, 177, 0), "solid"), 3)

        :Args (row-column):
            * **param1** (*int*): The row number (zero indexed).
            * **param2** (*int*): The column number (zero indexed).
            * **param3** (*str | List[str]*): Which side(s) of the cell to apply the border to.
            * **param4** (:py:class:`Border`): The border to add.
            * **param5** (*int*, *optinal*, default: 1): The length of the stroke to add.

        :Args (A1):
            * **param1** (*str*): A cell reference using Excel/Numbers-style A1 notation.
            * **param2** (*str | List[str]*): Which side(s) of the cell to apply the border to.
            * **param3** (:py:class:`Border`): The border to add.
            * **param4** (*int*, *optional*, default: 1): The length of the stroke to add.

        Raises
        ------
        TypeError:
            If an invalid number of arguments is passed or if the types of the arguments
            are invalid.

        Warns
        -----
        RuntimeWarning:
            If any of the sides to which the border is applied have been merged.
        """  # noqa: E501
        (row, col, *args) = self._validate_cell_coords(*args)
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
                self.set_cell_border(row, col, s, border_value, length)
            return

        cell = self._data[row][col]
        if cell.is_merged and (
            (side == "right" and cell.size[1] > 1) or (side == "bottom" and cell.size[0] > 1)
        ):
            warn(
                f"{side} edge of [{row},{col}] is merged; border not set",
                RuntimeWarning,
                stacklevel=2,
            )
            return
        elif isinstance(cell, MergedCell) and (
            (side == "top" and cell.row_start < row)
            or (side == "right" and cell.col_end > col)
            or (side == "bottom" and cell.row_end > row)
            or (side == "left" and cell.col_start < col)
        ):
            warn(
                f"{side} edge of [{row},{col}] is merged; border not set",
                RuntimeWarning,
                stacklevel=2,
            )
            return

        self._model.extract_strokes(self._table_id)

        if side in ["top", "bottom"]:
            for border_col_num in range(col, col + length):
                self._model.set_cell_border(self._table_id, row, border_col_num, side, border_value)
        elif side in ["left", "right"]:
            for border_row_num in range(row, row + length):
                self._model.set_cell_border(self._table_id, border_row_num, col, side, border_value)
        else:
            raise TypeError("side must be a valid border segment")

        self._model.add_stroke(self._table_id, row, col, side, border_value, length)

    def set_cell_formatting(self, *args: str, **kwargs) -> None:
        r"""
        Set the data format for a cell.

        Cell references can be **row-column** offsers or Excel/Numbers-style **A1** notation.

        .. code:: python

            table.set_cell_formatting(
                "C1",
                "date",
                date_time_format="EEEE, d MMMM yyyy"
            )
            table.set_cell_formatting(
                0,
                4,
                "number",
                decimal_places=3,
                negative_style=NegativeNumberStyle.RED
            )

        :Parameters:
            * **args** (*list*, *optional*) – Positional arguments for cell reference and data format type (see below)
            * **kwargs** (*dict*, *optional*) - Key-value pairs defining a formatting options for each data format (see below).

        :Args (row-column):
            * **param1** (*int*): The row number (zero indexed).
            * **param2** (*int*): The column number (zero indexed).
            * **param3** (*str*): Data format type for the cell (see "data formats" below).

        :Args (A1):
            * **param1** (*str*): A cell reference using Excel/Numbers-style A1 notation.
            * **param2** (*str*): Data format type for the cell (see "data formats" below).

        :Raises:
            * **TypeError** -
                If a tickbox is chosen for anything other than ``bool`` values.
            * **IndexError** -
                If the current cell value does not match a list of popup items.

        :Warns:
            * **RuntimeWarning** -
                If ``use_accounting_style`` is used with
                any ``negative_style`` other than ``NegativeNumberStyle.MINUS``, or
                if a rating is out of range 0 to 5 (rating is clamped to these values).

        All formatting styles share a name and a type, described in the **Common**
        parameters in the following table. Additional key-value pairs configure the format
        depending upon the value of ``kwargs["type"]``.

        :Common Args:
            * **name** (*str*) – The name of the custom format. If no name is provided,
              one is generated using the scheme ``Custom Format``, ``Custom Format 1``, ``Custom Format 2``, etc.
            * **type** (*str, optional, default: number*) – The type of format to
              create:

              * ``"base"``: A number base in the range 2-36.
              * ``"currency"``: A decimal formatted with a currency symbol.
              * ``"datetime"``: A date and time value with custom formatting.
              * ``"fraction"``: A number formatted as the nearest fraction.
              * ``"percentage"``: A number formatted as a percentage
              * ``"number"``: A decimal number.
              * ``"scientific"``: A decimal number with scientific notation.
              * ``"tickbox"``: A checkbox (bool values only).
              * ``"rating"``: A star rating from 0 to 5.
              * ``"slider"``: A range slider.
              * ``"stepper"``: An up/down value stepper.
              * ``"popup"``: A menu of options.

        :``"base"``:
            * **base_use_minus_sign** (*int, optional, default: 10*) – The integer
              base to represent the number from 2-36.
            * **base_use_minus_sign** (*bool, optional, default: True*) – If ``True``
              use a standard minus sign, otherwise format as two's compliment (only
              possible for binary, octal and hexadecimal.
            * **base_places** (*int, optional, default: 0*) – The number of
              decimal places, or ``None`` for automatic.

        :``"currency"``:
            * **currency** (*str, optional, default: "GBP"*) – An ISO currency
              code, e.g. ``"GBP"`` or ``"USD"``.
            * **decimal_places** (*int, optional, default: 2*) – The number of
              decimal places, or ``None`` for automatic.
            * **negative_style** (*:py:class:`~numbers_parser.NegativeNumberStyle`, optional, default: NegativeNumberStyle.MINUS*) – How negative numbers are represented.
              See `Negative number formats <#negative-formats>`_.
            * **show_thousands_separator** (*bool, optional, default: False*) – ``True``
              if the number should include a thousands seperator, e.g. ``,``
            * **use_accounting_style** (*bool, optional, default: False*) –  ``True``
              if the currency symbol should be formatted to the left of the cell and
              separated from the number value by a tab.

        :``"datetime"``:
            * **date_time_format** (*str, optional, default: "dd MMM YYY HH:MM"*) – A POSIX
               strftime-like formatting string of `Numbers date/time
               directives <#datetime-formats>`_.

        :``"fraction"``:
            * **fraction_accuracy** (*:py:class:`~numbers_parser.FractionAccuracy`, optional, default: FractionAccuracy.THREE* – The
                precision of the faction.

        :``"percentage"``:
            * **decimal_places** (*float, optional, default: None*) –  number of
              decimal places, or ``None`` for automatic.
            * **negative_style** (*:py:class:`~numbers_parser.NegativeNumberStyle`, optional, default: NegativeNumberStyle.MINUS*) – How negative numbers are represented.
              See `Negative number formats <#negative-formats>`_.
            * **show_thousands_separator** (*bool, optional, default: False*) – ``True``
              if the number should include a thousands seperator, e.g. ``,``

        :``"scientific"``:
            * **decimal_places** (*float, optional, default: None*) – number of
              decimal places, or ``None`` for automatic.

        :``"tickbox"``:
            * No additional parameters defined.

        :``"rating"``:
            * No additional parameters defined.

        :``"slider"``:
            * **control_format** (*ControlFormattingType, optional, default: ControlFormattingType.NUMBER*) - the format
                of the data in the slider. Valid options are ``"base"``, ``"currency"``,
                ``"datetime"``, ``"fraction"``, ``"percentage"``, ``"number"``,
                or ``"scientific". Each format allows additional parameters identical to those
                available for the formats themselves. For example, a slider using fractions
                is configured with ``fraction_accuracy``.
            * **increment** (*float, optional, default: 1*) - the slider's minimum value
            * **maximum** (*float, optional, default: 100*) - the slider's maximum value
            * **minimum** (*float, optional, default: 1*) - increment value for the slider

        :`"stepper"``:
            * **control_format** (*ControlFormattingType, optional, default: ControlFormattingType.NUMBER*) - the format
                of the data in the stepper. Valid options are ``"base"``, ``"currency"``,
                ``"datetime"``, ``"fraction"``, ``"percentage"``, ``"number"``,
                or ``"scientific"``. Each format allows additional parameters identical to those
                available for the formats themselves. For example, a stepper using fractions
                is configured with ``fraction_accuracy``.
            * **increment** (*float, optional, default: 1*) - the stepper's minimum value
            * **maximum** (*float, optional, default: 100*) - the stepper's maximum value
            * **minimum** (*float, optional, default: 1*) - increment value for the stepper

        :`"popup"``:
            * **popup_values** (*List[str|int|float], optional, default: None*) – values
                for the popup menu
            * **allow_none** (*bool, optional, default: True*) - If ``True``
                include a blank value in the list


        """  # noqa: E501
        (row, col, *args) = self._validate_cell_coords(*args)
        if len(args) == 1:
            format_type = args[0]
        elif len(args) > 1:
            raise TypeError("too many positional arguments to set_cell_formatting")
        else:
            raise TypeError("no type defined for cell format")

        if format_type == "custom":
            self._set_cell_custom_format(row, col, **kwargs)
        else:
            self._set_cell_data_format(row, col, format_type, **kwargs)

    def _set_cell_custom_format(self, row: int, col: int, **kwargs) -> None:
        if "format" not in kwargs:
            raise TypeError("no format provided for custom format")

        custom_format = kwargs["format"]
        if isinstance(custom_format, CustomFormatting):
            custom_format = kwargs["format"]
        elif isinstance(custom_format, str):
            if custom_format not in self._model.custom_formats:
                raise IndexError(f"format '{custom_format}' does not exist")
            custom_format = self._model.custom_formats[custom_format]
        else:
            raise TypeError("format must be a CustomFormatting object or format name")

        cell = self._data[row][col]
        type_name = type(cell).__name__
        format_type_name = custom_format.type.name.lower()
        if type_name not in CUSTOM_FORMATTING_ALLOWED_CELLS[format_type_name]:
            raise TypeError(
                f"cannot use {format_type_name} formatting for cells of type {type_name}"
            )

        format_id = self._model.custom_format_id(self._table_id, custom_format)
        cell._set_formatting(format_id, custom_format.type)

    def _set_cell_data_format(self, row: int, col: int, format_type_name: str, **kwargs) -> None:
        try:
            format_type = FormattingType[format_type_name.upper()]
        except (KeyError, AttributeError):
            raise TypeError(f"unsuported cell format type '{format_type_name}'") from None

        cell = self._data[row][col]
        type_name = type(cell).__name__
        if type_name not in FORMATTING_ALLOWED_CELLS[format_type_name]:
            raise TypeError(
                f"cannot use {format_type_name} formatting for cells of type {type_name}"
            )

        format = Formatting(type=format_type, **kwargs)
        if format_type_name in FORMATTING_ACTION_CELLS:
            control_id = self._model.control_cell_archive(self._table_id, format_type, format)
        else:
            control_id = None

        is_currency = True if format_type == FormattingType.CURRENCY else False
        if format_type_name in ["slider", "stepper"]:
            if "control_format" in kwargs:
                try:
                    control_format = kwargs["control_format"].name
                    number_format_type = FormattingType[control_format]
                    is_currency = (
                        True
                        if kwargs["control_format"] == ControlFormattingType.CURRENCY
                        else False
                    )
                except (KeyError, AttributeError):
                    control_format = kwargs["control_format"]
                    raise TypeError(
                        f"unsupported number format '{control_format}' for {format_type_name}"
                    ) from None
            else:
                number_format_type = FormattingType.NUMBER
            format_id = self._model.format_archive(self._table_id, number_format_type, format)
        elif format_type_name == "popup":
            if cell.value not in format.popup_values:
                raise IndexError(
                    f"current cell value '{cell.value}' does not match any popup values"
                )

            popup_format_type = FormattingType.TEXT if isinstance(cell, TextCell) else True
            format_id = self._model.format_archive(self._table_id, popup_format_type, format)
        else:
            format_id = self._model.format_archive(self._table_id, format_type, format)

        cell._set_formatting(format_id, format_type, control_id, is_currency=is_currency)
