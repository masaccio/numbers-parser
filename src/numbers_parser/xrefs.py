"""Parse and resolve cell references and named ranges."""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field, replace
from enum import IntEnum
from typing import TYPE_CHECKING, Any, cast

from numbers_parser.constants import OPERATOR_PRECEDENCE

if TYPE_CHECKING:
    from collections.abc import Callable, Mapping, Sequence

    from numbers_parser.model import NumbersModel

__all__ = [
    "xl_cell_to_rowcol",
    "xl_col_to_name",
    "xl_col_to_offset",
    "xl_range",
    "xl_rowcol_to_cell",
]


class TableAxis(IntEnum):
    """Indicates whether a cache is for rows or columns."""

    ROW = 1
    COLUMN = 2


class RefScope(IntEnum):
    """The required scope for a name reference in a document."""

    # Name is unique to the document.
    DOCUMENT = 1
    # Name is unique to a sheet.
    SHEET = 2
    # Name is unique to a table and that table is uniquely named.
    TABLE = 3
    # All other names.
    NONE = 4


@dataclass
class ScopedNameRef:
    """Represent a named reference and the scope where it is unique."""

    name: str
    offset: int | None = None
    axis: TableAxis | None = None
    table_id: int | None = None
    scope: RefScope = RefScope.NONE


class CellRangeType(IntEnum):
    """Enumerate supported cell-reference range forms."""

    ROW_RANGE = 1
    COL_RANGE = 2
    RANGE = 3
    NAMED_RANGE = 4
    CELL = 5
    NAMED_ROW_COLUMN = 6


@dataclass
class CellRange:
    """Parse and format a cell or named-reference range."""

    _model: NumbersModel

    row_start: int | None = None
    row_end: int | None = None
    col_start: int | None = None
    col_end: int | None = None
    row_start_is_abs: bool = False
    row_end_is_abs: bool = False
    col_start_is_abs: bool = False
    col_end_is_abs: bool = False
    from_table_id: int = 0
    to_table_id: int = 0
    _ref_error: bool = False
    _do_init: bool = True
    _table_names: list[str] | None = field(init=False, default=None, repr=False)

    def __post_init__(self) -> None:
        """Validate and initialize the instance state."""
        if (self.row_start is not None and self.row_start < 0) or (
            self.col_start is not None and self.col_start < 0
        ):
            self._ref_error = True
        else:
            self._initialize_table_data()
            self._model.name_ref_cache.refresh()
            self._set_sheet_ids()

    def _initialize_table_data(self) -> None:
        """Initialize table data."""
        self._table_names = self._model.table_names()
        self.table_name_unique = {
            name: self._table_names.count(name) == 1 for name in self._table_names
        }

    def _set_sheet_ids(self) -> None:
        """Determine the sheet IDs for the referenced tables."""
        if self.to_table_id is None:
            self.to_table_id = self.from_table_id
        self.from_sheet_id = self._model.table_id_to_sheet_id(self.from_table_id)
        self.to_sheet_id = self._model.table_id_to_sheet_id(self.to_table_id)

    def expand_ref(
        self,
        ref: str | ScopedNameRef,
        is_abs: bool = False,
        no_prefix: bool = False,
    ) -> str:
        """Expand ref."""
        if self._ref_error:
            return "#REF!"

        self._model.name_ref_cache.refresh()
        is_document_unique = (
            ref.scope == RefScope.DOCUMENT if isinstance(ref, ScopedNameRef) else False
        )
        is_sheet_unique = ref.scope == RefScope.SHEET if isinstance(ref, ScopedNameRef) else False
        is_table_unique = ref.scope == RefScope.TABLE if isinstance(ref, ScopedNameRef) else False

        if isinstance(ref, ScopedNameRef):
            ref_str = f"${ref.name}" if is_abs else ref.name
        else:
            ref_str = f"${ref}" if is_abs else ref
        if any(x in ref_str for x in OPERATOR_PRECEDENCE):
            ref_str = f"'{ref_str}'"
        elif "'" in ref_str:
            ref_str = ref_str.replace("'", "'''")

        if no_prefix or is_document_unique:
            return ref_str

        if self.from_table_id == self.to_table_id:
            return ref_str

        table_name = self._model.table_name(self.to_table_id)
        if self.from_sheet_id == self.to_sheet_id and is_sheet_unique:
            # If absolute Numbers seems to unnecessarily include the table name
            return f"{table_name}::{ref_str}" if is_abs else ref_str

        is_table_unique |= self.table_name_unique[table_name]
        if self.from_sheet_id == self.to_sheet_id or is_table_unique:
            return f"{table_name}::{ref_str}"

        sheet_name = self._model.sheet_name(self.to_sheet_id)
        return f"{sheet_name}::{table_name}::{ref_str}"

    def __str__(self) -> str:
        """Return the string representation."""
        if self._ref_error:
            return "#REF!"

        self._model.name_ref_cache.refresh()
        # Handle row-only ranges
        if self.col_start is None:
            self.row_start = cast("int", self.row_start)
            row_range = self._model.name_ref_cache.row_ranges[self.to_table_id]
            return self._format_row_range(self.row_start, self.row_end, row_range)

        # Handle column-only ranges
        if self.row_start is None:
            self.col_start = cast("int", self.col_start)
            col_range = self._model.name_ref_cache.col_ranges[self.to_table_id]
            return self._format_col_range(self.col_start, self.col_end, col_range)

        # Handle full cell ranges
        return self._format_cell_range(self.row_start, self.col_start, self.row_end, self.col_end)

    def _format_row_range(
        self,
        row_start: int,
        row_end: int | None,
        row_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a row-only range."""
        if row_end is None:
            return self._format_single_row(row_start, row_range)
        return self._format_row_span(row_start, row_end, row_range)

    def _format_single_row(
        self,
        row_start: int,
        row_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a single row, either numeric or named."""
        ref = row_range.get(row_start)
        if ref is None:
            return self._format_numeric_row(row_start)
        return self.expand_ref(ref, self.row_start_is_abs)

    def _format_numeric_row(self, row_start: int) -> str:
        """Format a single numeric row."""
        return ":".join(
            [
                self.expand_ref(str(row_start + 1), self.row_start_is_abs),
                self.expand_ref(str(row_start + 1), self.row_start_is_abs, no_prefix=True),
            ],
        )

    def _format_row_span(
        self,
        row_start: int,
        row_end: int,
        row_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a range of rows."""
        start_ref = row_range.get(row_start)
        end_ref = row_range.get(row_end)

        if start_ref is None or end_ref is None:
            return ":".join(
                [
                    self.expand_ref(str(row_start + 1), self.row_start_is_abs),
                    self.expand_ref(str(row_end + 1), self.row_end_is_abs, no_prefix=True),
                ],
            )
        return ":".join(
            [
                self.expand_ref(
                    start_ref,
                    self.row_start_is_abs,
                    no_prefix=RefScope.DOCUMENT in (start_ref.scope, end_ref.scope),
                ),
                self.expand_ref(end_ref, self.row_end_is_abs, no_prefix=True),
            ],
        )

    def _format_col_range(
        self,
        col_start: int,
        col_end: int | None,
        col_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a column-only range."""
        if col_end is None:
            return self._format_single_column(col_start, col_range)
        return self._format_column_span(col_start, col_end, col_range)

    def _format_single_column(
        self,
        col_start: int,
        col_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a single column, either numeric or named."""
        ref = col_range.get(col_start)
        if ref is None:
            return self.expand_ref(xl_col_to_name(col_start, col_abs=self.col_start_is_abs))
        return self.expand_ref(ref, self.col_start_is_abs)

    def _format_column_span(
        self,
        col_start: int,
        col_end: int,
        col_range: Mapping[int, ScopedNameRef | None],
    ) -> str:
        """Format a range of columns."""
        start_ref = col_range.get(col_start)
        end_ref = col_range.get(col_end)

        if start_ref is None or end_ref is None:
            start_name = self.expand_ref(
                xl_col_to_name(col_start, col_abs=self.col_start_is_abs),
            )
            end_name = self.expand_ref(
                xl_col_to_name(col_end, col_abs=self.col_end_is_abs),
                no_prefix=True,
            )
            return f"{start_name}:{end_name}"
        return ":".join(
            [
                self.expand_ref(
                    start_ref,
                    self.col_start_is_abs,
                    no_prefix=RefScope.DOCUMENT in (start_ref.scope, end_ref.scope),
                ),
                self.expand_ref(end_ref, self.col_end_is_abs, no_prefix=True),
            ],
        )

    def _format_cell_range(
        self,
        row_start: int,
        col_start: int,
        row_end: int | None,
        col_end: int | None,
    ) -> str:
        """Format a full cell range."""
        if row_end is None or col_end is None:
            return self.expand_ref(
                xl_rowcol_to_cell(
                    row_start,
                    col_start,
                    row_abs=self.row_start_is_abs,
                    col_abs=self.col_start_is_abs,
                ),
            )
        return ":".join(
            [
                self.expand_ref(
                    xl_rowcol_to_cell(
                        row_start,
                        col_start,
                        row_abs=self.row_start_is_abs,
                        col_abs=self.col_start_is_abs,
                    ),
                ),
                self.expand_ref(
                    xl_rowcol_to_cell(
                        row_end,
                        col_end,
                        row_abs=self.row_end_is_abs,
                        col_abs=self.col_end_is_abs,
                    ),
                    no_prefix=True,
                ),
            ],
        )


class ScopedNameRefCache:
    """Cache named row, column, and table references."""

    def __init__(self, model: NumbersModel) -> None:
        """Initialize the instance state."""
        self._model = model
        self.doc_name_refs: dict[str, ScopedNameRef | None] = {}
        self.sheet_name_refs: dict[int, dict[str, ScopedNameRef | None]] = {}
        self.table_name_refs: dict[int, dict[str, ScopedNameRef | None]] = {}
        self.row_ranges: dict[int, dict[int, ScopedNameRef | None]] = {}
        self.col_ranges: dict[int, dict[int, ScopedNameRef | None]] = {}
        self._dirty_cache = True
        self.table_names: list[str] = []

    def mark_dirty(self) -> None:
        """Mark dirty."""
        self._dirty_cache = True

    def refresh(self) -> None:
        """Refresh the value."""
        if self._dirty_cache:
            self.calculate_named_ranges()
            self._dirty_cache = False

    @staticmethod
    def _exact_count(pool: Sequence[Any], value: str) -> int:
        """Return the number of exact references to a name."""
        return sum(1 for x in pool if type(x) is type(value) and x == value)

    def _row_data(self, table_id: int, row: int) -> int | str | bool | None:
        """Collect row reference names and their scopes."""
        num_header_cols = self._model.num_header_cols(table_id)
        return self._model._table_data[table_id][row][num_header_cols - 1].formatted_value  # noqa: SLF001

    def _column_data(self, table_id: int, col: int) -> int | str | bool | None:
        """Collect column reference names and their scopes."""
        num_header_rows = self._model.num_header_rows(table_id)
        return self._model._table_data[table_id][num_header_rows - 1][col].formatted_value  # noqa: SLF001

    def _calculate_name_scopes(
        self,
        sheet_id: int,
        table_id: int,
        axis: TableAxis,
        doc_name_counts: dict[str, int],
        sheet_name_counts: dict[int, dict[str, int]],
    ) -> dict[int, str | None]:
        """Calculate name scopes."""
        num_header_rows = self._model.num_header_rows(table_id)
        num_header_cols = self._model.num_header_cols(table_id)

        data_lookup: Callable[[int, int], int | str | bool | None]
        if axis == TableAxis.ROW:
            data_lookup = self._row_data
            range_start = num_header_rows
            range_end = self._model.number_of_rows(table_id)
            first_offset = num_header_rows
        else:
            data_lookup = self._column_data
            range_start = num_header_cols
            range_end = self._model.number_of_columns(table_id)
            first_offset = num_header_cols

        if (axis == TableAxis.ROW and num_header_cols == 0) or (
            axis == TableAxis.COLUMN and num_header_rows == 0
        ):
            return dict.fromkeys(range(range_end))

        scopes: dict[int, str | None] = {}
        names: list[str | None] = []
        all_names = [data_lookup(table_id, idx) for idx in range(range_start, range_end)]

        for idx in range(range_end):
            if idx < first_offset:
                names.append(None)
                scopes[idx] = None
            else:
                raw_name = data_lookup(table_id, idx)
                if self._exact_count(all_names, str(raw_name)) > 1 or raw_name is None:
                    names.append(None)
                    scopes[idx] = None
                else:
                    name_str = str(raw_name)
                    names.append(name_str)
                    scopes[idx] = name_str

        for parsed_name in names:
            if parsed_name is not None:
                doc_name_counts[parsed_name] += 1
                sheet_name_counts[sheet_id][parsed_name] += 1

        return scopes

    def _calculate_scope_types(
        self,
        sheet_id: int,
        table_id: int,
        axis: TableAxis,
        scopes: dict[int, str | None],
    ) -> dict[int, ScopedNameRef | None]:
        """
        Classify locally unique row and column names by their scope.

        A name may be unique to a table, sheet, or document.
        """
        new_scopes: dict[int, ScopedNameRef | None] = {}
        table_name = self._model.table_name(table_id)

        if axis == TableAxis.ROW:
            axis_range = range(self._model.number_of_rows(table_id))
        else:
            axis_range = range(self._model.number_of_columns(table_id))

        for idx in axis_range:
            name = scopes.get(idx)
            if name is None:
                new_scopes[idx] = None
                continue

            scope = ScopedNameRef(name, axis=axis, table_id=table_id, offset=idx)
            if name in self.doc_name_refs:
                new_scopes[idx] = replace(scope, scope=RefScope.DOCUMENT)
                self.doc_name_refs[name] = new_scopes[idx]
            elif name in self.sheet_name_refs[sheet_id]:
                new_scopes[idx] = replace(scope, scope=RefScope.SHEET)
                self.sheet_name_refs[sheet_id][name] = new_scopes[idx]
            else:
                scope_type = (
                    RefScope.TABLE
                    if self._exact_count(self.table_names, table_name) == 1
                    else RefScope.NONE
                )
                new_scopes[idx] = replace(scope, scope=scope_type)
                self.table_name_refs[table_id][name] = new_scopes[idx]

        return new_scopes

    def _calculate_table_name_maps(self) -> None:
        """Calculate table name maps."""
        self.sheet_name_to_id = {
            self._model.sheet_name(sid): sid for sid in self._model.sheet_ids()
        }
        self.sheet_id_to_name = {
            sid: self._model.sheet_name(sid) for sid in self._model.sheet_ids()
        }
        self.unique_table_name_to_id = {
            self._model.table_name(tid): tid
            for tid in self._model.table_ids()
            if self.table_names.count(self._model.table_name(tid)) == 1
        }
        self.sheet_table_name_to_id = {
            self._model.sheet_name(sid): {
                self._model.table_name(tid): tid for tid in self._model.table_ids(sid)
            }
            for sid in self._model.sheet_ids()
        }

    def calculate_named_ranges(self) -> None:
        """
        Find unique row and column headers for range references.

        Record document-wide and table-specific names by table ID.
        """
        self.table_names = self._model.table_names()
        self._calculate_table_name_maps()

        doc_name_counts: dict[str, int] = defaultdict(int)
        sheet_name_counts: dict[int, dict[str, int]] = {}
        row_ranges_str: dict[int, dict[int, str | None]] = {}
        col_ranges_str: dict[int, dict[int, str | None]] = {}

        for sheet_id in self._model.sheet_ids():
            sheet_name_counts[sheet_id] = defaultdict(int)
            for table_id in self._model.table_ids(sheet_id):
                self.table_name_refs[table_id] = {}
                row_ranges_str[table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                    doc_name_counts,
                    sheet_name_counts,
                )
                col_ranges_str[table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                    doc_name_counts,
                    sheet_name_counts,
                )

        # Init the list of document-scoped names if they are unique
        self.doc_name_refs = {name: None for name, count in doc_name_counts.items() if count == 1}

        self.sheet_name_refs = {}
        for sheet_id in self._model.sheet_ids():
            # Init the list of sheet-scoped names if they are unique
            self.sheet_name_refs[sheet_id] = {
                name: None for name, count in sheet_name_counts[sheet_id].items() if count == 1
            }

            for table_id in self._model.table_ids(sheet_id):
                self.row_ranges[table_id] = self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                    row_ranges_str[table_id],
                )
                self.col_ranges[table_id] = self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                    col_ranges_str[table_id],
                )


# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")

col_parts = re.compile(r"^(\$?)([A-Z]{1,3})$")


def xl_col_to_offset(col_str: str) -> int:
    """
    Convert a column reference in A1 notation to a zero indexed column.

    Parameters
    ----------
    col_str:  str
        A1 notation column reference

    Returns
    -------
    col: int
        Column numbers (zero indexed).

    """
    if not col_str:
        return 0

    match = col_parts.match(col_str)
    if not match:
        msg = f"invalid cell reference {col_str}"
        raise IndexError(msg)

    col_str = match.group(2)

    # Convert base26 column string to number.
    col = 0
    for expn, char in enumerate(reversed(match.group(2))):
        col += (ord(char) - ord("A") + 1) * (26**expn)

    # Convert 1-index to zero-index
    col -= 1

    return col


def xl_cell_to_rowcol(cell_str: str) -> tuple[int, int]:
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

    Parameters
    ----------
    cell_str:  str
        A1 notation cell reference
        An empty string is treated as the first cell (A1).

    Returns
    -------
    row, col: int, int
        Cell row and column numbers (zero indexed).

    """
    if not cell_str:
        return 0, 0

    match = range_parts.match(cell_str)
    if not match:
        msg = f"invalid cell reference {cell_str}"
        raise IndexError(msg)

    col_str = match.group(2)
    row_str = match.group(4)

    if int(row_str) < 1:
        msg = f"invalid cell reference {cell_str}"
        raise IndexError(msg)

    # Convert base26 column string to number.
    col = 0
    for expn, char in enumerate(reversed(col_str)):
        col += (ord(char) - ord("A") + 1) * (26**expn)

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col


def xl_range(first_row: int, first_col: int, last_row: int, last_col: int) -> str:
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
        A1:B1 style range string, or a single-cell reference if both coordinates
        identify the same cell.

    """
    range1 = xl_rowcol_to_cell(first_row, first_col)
    range2 = xl_rowcol_to_cell(last_row, last_col)

    if range1 == range2:
        return range1
    return range1 + ":" + range2


def xl_rowcol_to_cell(row: int, col: int, row_abs: bool = False, col_abs: bool = False) -> str:
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
        msg = f"row reference {row} below zero"
        raise IndexError(msg)

    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    row += 1  # Change to 1-index.
    row_abs_str = "$" if row_abs else ""

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs_str + str(row)


def xl_col_to_name(col: int, col_abs: bool = False) -> str:
    """
    Convert a zero indexed column cell reference to a string.

    Parameters
    ----------
    col: int
        The column number (zero indexed).
    col_abs: bool, default: False
        If ``True``, make the column absolute.

    Returns
    -------
        str:
            Column in A1 notation.

    """
    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    col += 1  # Change to 1-index.
    col_str = ""
    col_abs_str = "$" if col_abs else ""

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

    return col_abs_str + col_str
