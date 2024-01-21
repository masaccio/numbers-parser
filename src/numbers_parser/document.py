from typing import Dict, Generator, List, Tuple, Union
from warnings import warn

from numbers_parser.cell import (
    Border,
    Cell,
    CustomFormatting,
    CustomFormattingType,
    DateCell,
    Formatting,
    FormattingType,
    MergedCell,
    NumberCell,
    Style,
    TextCell,
    xl_cell_to_rowcol,
)
from numbers_parser.cell_storage import CellStorage
from numbers_parser.constants import (
    DEFAULT_COLUMN_COUNT,
    DEFAULT_ROW_COUNT,
    MAX_COL_COUNT,
    MAX_HEADER_COUNT,
    MAX_ROW_COUNT,
)
from numbers_parser.containers import ItemsList
from numbers_parser.file import write_numbers_file
from numbers_parser.model import _NumbersModel
from numbers_parser.numbers_cache import Cacheable, cache


class Sheet:
    pass


class Table:
    pass


"""Add a new sheet to the current document. If no sheet name is provided,
the next available numbered sheet will be generated.

:param sheet_name: the name of the sheet to add to the document
    If ``sheet_name`` is ``None``, the next available sheet name in the
    series ``Sheet 1``, ``Sheet 2``, etc. is chosen.
:param table_name: the name of the table created in the new sheet, defaults to ``Table 1``
:type table_name: str, optional
:param num_rows: the number of columns in the newly created table
:type num_rows: int, optional
:param num_cols: the number of columns in the newly created table
:type num_cols: int, optional
:raises IndexError: if the sheet name already exists in the document.
"""


class Document:
    """Create an instance of a new Numbers document. If no document filename
    is provided, a new document is created with the parameters passed to the
    ``Document`` constructor.

    :param filename: Apple Numbers document to read.
        If ``filename`` is ``None``, an empty document is created using the defaults
        defined by the class constructor. You can optionionally override these
        defaults at object construction time.
    :type filename: str, optional
    :param sheet_name: name of the first sheet in a new document
    :type sheet_name: str, optional
    :param table_name: name of the first table in the first sheet of a new
    :type table_name: str, optional
    :param num_header_rows: number of header rows in the first table of a new document.
    :type: num_header_rows: int, optional
    :param num_header_cols: number of header columns in the first table of a new document.
    :type: num_header_cols: int, optional
    :param num_rows: number of rows in the first table of a new document.
    :type: num_rows: int, optional
    :param num_cols: number of columns in the first table of a new document.
    :type: num_cols: int, optional
    :raises IndexError: if the sheet name already exists in the document.
    :raises IndexError: if the table name already exists in the first sheet.

    Examples

    Reading a document and examining the ``Tables`` object:

    .. code-block:: python

        >>> from numbers_parser import Document
        >>> doc = Document("mydoc.numbers")
        >>> doc.sheets[0].name
        'Sheet 1'
        >>> table = doc.sheets[0].tables[0]
        >>> table.name
        'Table 1'

    Creating a new document:

    .. code-block:: python

        doc = Document()
        doc.add_sheet("New Sheet", "New Table")
        sheet = doc.sheets["New Sheet"]
        table = sheet.tables["New Table"]
        table.write(1, 1, 1000)
        table.write(1, 2, 2000)
        table.write(1, 3, 3000)
        doc.save("mydoc.numbers")
    """

    def __init__(  # noqa: PLR0913
        self,
        filename: str = None,
        sheet_name: str = "Sheet 1",
        table_name: str = "Table 1",
        num_header_rows: int = 1,
        num_header_cols: int = 1,
        num_rows: int = DEFAULT_ROW_COUNT,
        num_cols: int = DEFAULT_COLUMN_COUNT,
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
        """
        Return a list of all sheets in the document.

        :return: List of :class:`Sheet` objects in the document
        :rtype: List[:class:`Sheet`]
        """
        return self._sheets

    @property
    def styles(self) -> Dict[str, Style]:
        """
        Return a dict of styles available in the document.

        :return: Dict of :class:`Style` objects in the document with the
            style name as keys.
        :rtype: Dict[:class:`Style`]
        """
        return self._model.styles

    @property
    def custom_formats(self) -> Dict[str, CustomFormatting]:
        """
        Return a dict of custom formats available in the document.

        :return: Dict of :class:`CustomFormatting` objects in the document with the
            format name as keys.
        :rtype: Dict[:class:`CustomFormatting`]
        """
        return self._model.custom_formats

    def save(self, filename: str) -> None:
        """Save the document in the specified filename

        :param filename: the path to save the document to
            If the file already exists, it will be overwritten.
        """
        for sheet in self.sheets:
            for table in sheet.tables:
                self._model.recalculate_table_data(table._table_id, table._data)
        write_numbers_file(filename, self._model.file_store)

    def add_sheet(
        self,
        sheet_name: str = None,
        table_name: str = "Table 1",
        num_rows: int = DEFAULT_ROW_COUNT,
        num_cols: int = DEFAULT_COLUMN_COUNT,
    ) -> None:
        """Add a new sheet to the current document. If no sheet name is provided,
        the next available numbered sheet will be generated.

        :param sheet_name: the name of the sheet to add to the document
            If ``sheet_name`` is ``None``, the next available sheet name in the
            series ``Sheet 1``, ``Sheet 2``, etc. is chosen.
        :type sheet_name: str, optional
        :param table_name: the name of the table created in the new sheet, defaults to ``Table 1``
        :type table_name: str, optional
        :param num_rows: the number of columns in the newly created table
        :type num_rows: int, optional
        :param num_cols: the number of columns in the newly created table
        :type num_cols: int, optional
        :raises IndexError: if the sheet name already exists in the document.
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
        r"""Add a new style to the current document. If no style name is
        provided, the next available numbered style will be generated.

        :param \**kwargs: style arguments
            Key-value pairs defining a cell style (see below)

        :Style Keyword Arguments:
            * *alignment* (**Alignment**): horizontal and vertical alignment of the cell
            * *bg_color* (**Union[RGB, List[RGB]]**): cell background color or list
              of colors for gradients
            * *bold* (**str**) : ``True`` if the cell font is bold
            * *font_color* (**RGB**) : font color
            * *font_size* (**float**) : font size in points
            * *font_name* (**str**) : font name
            * *italic* (**str**) : ``True`` if the cell font is italic
            * *name* (**str**) : cell style
            * *underline* (**str**) : ``True`` if the cell font is underline
            * *strikethrough* (**str**) : ``True`` if the cell font is strikethrough
            * *first_indent* (**float**) : first line indent in points
            * *left_indent* (**float**) : left indent in points
            * *right_indent* (**float**) : right indent in points
            * *text_inset* (**float**) : text inset in points
            * *text_wrap* (**str**) : ``True`` if text wrapping is enabled

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

        """
        if "name" in kwargs and kwargs["name"] is not None and kwargs["name"] in self._model.styles:
            raise IndexError(f"style '{kwargs['name']}' already exists")
        style = Style(**kwargs)
        if style.name is None:
            style.name = self._model.custom_style_name()
        style._update_styles = True
        self._model.styles[style.name] = style
        return style

    def add_custom_format(self, **kwargs) -> CustomFormatting:
        r"""Add a new custom format to the current document. If no format name is
        provided, the next available numbered format will be generated.

        :param \**kwargs: style arguments
            Key-value pairs defining a cell format (see below)

        :Common Custom Formatting Keyword Arguments:
            * *alignment* (**Alignment**): the horizontal and vertical alignment of the cell
            * *name* (**str**): the name of the custom format
              If no name is provided, one is generated using the scheme ``Custom Format``,
              ``Custom Format 1``, ``Custom Format 2``, etc.
            * *type* (**str**): the type of format to create
              Supported formats are ``number``, ``datetime`` and ``text``. If no type is
              provided, ``add_custom_format`` defaults to ``number``

        :Custom Formatting Keyword Arguments for ``type``=``number``:
            * *integer_format* (**PaddingType**): how to pad integers, default ``PaddingType.NONE``
            * *decimal_format* (**PaddingType**): how to pad decimals, default ``PaddingType.NONE``
            * *num_integers* (**int**): integer precision when integers are padded, default 0
            * *num_decimals* (**int**): integer precision when decimals are padded, default 0
            * *show_thousands_separator* (**bool**): ``True`` if the number should include
              a thousands seperator

        :Custom Formatting Keyword Arguments for ``type``=``datetime``:
            * *format* (**str**): a POSIX strftime-like formatting string
              See Date/time formatting for a list of supported directives, default ``d MMM y``

        :Custom Formatting Keyword Arguments for ``type``=``text``:
            * *format* (**str**): string format
              The cell value is inserted in place of %s. Only one substitution is allowed by
              Numbers, and multiple %s formatting references raise a TypeError exception
        """
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
        elif custom_format.type == CustomFormattingType.TEXT:
            self._model.add_custom_text_format_archive(custom_format)
        return custom_format


class Sheet:
    """Do not instantiate directly. Sheets are created by ``Document``"""

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


class Table(Cacheable):  # noqa: F811
    """Do not instantiate directly. Tables are created by ``Document``"""

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

    def write(self, *args, style=None):
        # TODO: write needs to retain/init the border
        (row_num, col_num, value) = self._validate_cell_coords(*args)
        self._data[row_num][col_num] = Cell.from_value(row_num, col_num, value)
        storage = CellStorage(self._model, self._table_id, None, row_num, col_num)
        storage.update_value(value, self._data[row_num][col_num])
        self._data[row_num][col_num].update_storage(storage)

        merge_cells = self._model.merge_cells(self._table_id)
        self._data[row_num][col_num]._table_id = self._table_id
        self._data[row_num][col_num]._model = self._model
        self._data[row_num][col_num]._set_merge(merge_cells.get((row_num, col_num)))

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

    def set_cell_formatting(self, *args: str, **kwargs) -> None:
        """Set the formatting for a cell."""
        (row_num, col_num, *args) = self._validate_cell_coords(*args)
        if len(args) == 1:
            format_type = args[0]
        elif len(args) > 1:
            raise TypeError("too many positional arguments to set_cell_formatting")
        else:
            raise TypeError("no type defined for cell format")

        if format_type == "custom":
            self.set_cell_custom_format(row_num, col_num, **kwargs)
        else:
            self.set_cell_data_format(row_num, col_num, format_type, **kwargs)

    def set_cell_custom_format(self, row_num: int, col_num: int, **kwargs) -> None:
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

        cell = self._data[row_num][col_num]
        if custom_format.type == CustomFormattingType.DATETIME and not isinstance(cell, DateCell):
            type_name = type(cell).__name__
            raise TypeError(f"cannot use date/time formatting for cells of type {type_name}")
        elif custom_format.type == CustomFormattingType.NUMBER and not isinstance(cell, NumberCell):
            type_name = type(cell).__name__
            raise TypeError(f"cannot use date/time formatting for cells of type {type_name}")
        elif custom_format.type == CustomFormattingType.TEXT and not isinstance(cell, TextCell):
            type_name = type(cell).__name__
            raise TypeError(f"cannot set formatting for cells of type {type_name}")

        if custom_format.type == CustomFormattingType.NUMBER:
            format_id = self._model.custom_decimal_format_id(self._table_id, custom_format)
        elif custom_format.type == CustomFormattingType.TEXT:
            format_id = self._model.custom_text_format_id(self._table_id, custom_format)
        cell._set_formatting(format_id, custom_format.type)

    def set_cell_data_format(self, row_num: int, col_num: int, format_type: str, **kwargs) -> None:
        try:
            format_type = FormattingType[format_type.upper()]
        except (KeyError, AttributeError):
            raise TypeError(f"unsuported cell format type '{format_type}'") from None

        cell = self._data[row_num][col_num]
        if format_type == FormattingType.DATETIME and not isinstance(cell, DateCell):
            type_name = type(cell).__name__
            raise TypeError(f"cannot use date/time formatting for cells of type {type_name}")
        elif not isinstance(cell, NumberCell) and not isinstance(cell, DateCell):
            type_name = type(cell).__name__
            raise TypeError(f"cannot set formatting for cells of type {type_name}")

        format = Formatting(type=format_type, **kwargs)
        format_id = self._model.format_archive(self._table_id, format_type, format)
        cell._set_formatting(format_id, format_type)
