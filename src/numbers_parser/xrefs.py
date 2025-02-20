from __future__ import annotations

import re
from collections import defaultdict
from contextlib import suppress
from dataclasses import dataclass, field, replace
from enum import IntEnum

from numbers_parser.constants import OPERATOR_PRECEDENCE

__all__ = ["xl_cell_to_rowcol", "xl_col_to_name", "xl_range", "xl_rowcol_to_cell"]


class TableAxis(IntEnum):
    """Indicates whether a cache is for rows or columns."""

    ROW = 1
    COLUMN = 2


class RefScope(IntEnum):
    """The required scope for a name reference in a document."""

    """Name is unique to the document."""
    DOCUMENT = 1
    """Name is unique to a sheet."""
    SHEET = 2
    """Name is unique to a table and that table is uniquely named."""
    TABLE = 3
    """All other names."""
    NONE = 4


@dataclass
class ScopedNameRef:
    name: str
    offset: int = None
    axis: TableAxis = None
    table_id: int = None
    scope: RefScope = RefScope.NONE


class CellRangeType(IntEnum):
    ROW_RANGE = 1
    COL_RANGE = 2
    RANGE = 3
    NAMED_RANGE = 4
    CELL = 5
    NAMED_ROW_COLUMN = 6


@dataclass
class CellRange:
    model: object = None
    row_start: int | str = None
    row_end: int = None
    col_start: int | str = None
    col_end: int = None
    row_start_is_abs: bool = False
    row_end_is_abs: bool = False
    col_start_is_abs: bool = False
    col_end_is_abs: bool = False
    from_table_id: int = None
    to_table_id: int = None
    range_type: CellRangeType = None
    _table_names: list[str] = field(init=False, default=None, repr=False)

    def __post_init__(self):
        if self._table_names is None:
            self._initialize_table_data()
        if len(self.model.name_ref_cache.row_ranges) == 0:
            self.model.name_ref_cache.calculate_named_ranges()
        self.name_ref_cache = self.model.name_ref_cache
        self._set_sheet_ids()

    def _initialize_table_data(self):
        self._table_names = self.model.table_names()
        self.table_name_unique = {
            name: self._table_names.count(name) == 1 for name in self._table_names
        }

    def _set_sheet_ids(self):
        """Determine the sheet IDs for the referenced tables."""
        if self.to_table_id is None:
            self.to_table_id = self.from_table_id
        self.from_sheet_id = self.model.table_id_to_sheet_id(self.from_table_id)
        self.to_sheet_id = self.model.table_id_to_sheet_id(self.to_table_id)

    def expand_ref(self, ref: str, is_abs: bool = False, no_prefix=False) -> str:
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

        table_name = self.model.table_name(self.to_table_id)
        if self.from_sheet_id == self.to_sheet_id and is_sheet_unique:
            # If absolute Numbers seems to unnecessarily include the table name
            return f"{table_name}::{ref_str}" if is_abs else ref_str

        is_table_unique |= self.table_name_unique[table_name]
        if self.from_sheet_id == self.to_sheet_id or is_table_unique:
            return f"{table_name}::{ref_str}"

        sheet_name = self.model.sheet_name(self.to_sheet_id)
        return f"{sheet_name}::{table_name}::{ref_str}"

    def __str__(self):
        # Handle row-only ranges
        if self.col_start is None:
            row_range = self.name_ref_cache.row_ranges[self.to_table_id]
            return self._format_row_range(self.row_start, self.row_end, row_range)

        # Handle column-only ranges
        if self.row_start is None:
            col_range = self.name_ref_cache.col_ranges[self.to_table_id]
            return self._format_col_range(self.col_start, self.col_end, col_range)

        # Handle full cell ranges
        return self._format_cell_range(self.row_start, self.col_start, self.row_end, self.col_end)

    def _format_row_range(self, row_start, row_end, row_range):
        """Formats a row-only range."""
        if row_end is None:
            return self._format_single_row(row_start, row_range)
        return self._format_row_span(row_start, row_end, row_range)

    def _format_single_row(self, row_start, row_range):
        """Formats a single row, either numeric or named."""
        if row_range[row_start] is None:
            return self._format_numeric_row(row_start)
        return self.expand_ref(row_range[row_start], self.row_start_is_abs)

    def _format_numeric_row(self, row_start):
        """Formats a single numeric row."""
        return ":".join(
            [
                self.expand_ref(str(row_start + 1), self.row_start_is_abs),
                self.expand_ref(str(row_start + 1), self.row_start_is_abs, no_prefix=True),
            ],
        )

    def _format_row_span(self, row_start, row_end, row_range):
        """Formats a range of rows."""
        if row_range[row_start] is None:
            return ":".join(
                [
                    self.expand_ref(str(row_start + 1), self.row_start_is_abs),
                    self.expand_ref(str(row_end + 1), self.row_end_is_abs, no_prefix=True),
                ],
            )
        return ":".join(
            [
                self.expand_ref(
                    row_range[row_start],
                    self.row_start_is_abs,
                    no_prefix=row_range[row_start].scope == RefScope.DOCUMENT
                    or row_range[row_end].scope == RefScope.DOCUMENT,
                ),
                self.expand_ref(row_range[row_end], self.row_end_is_abs, no_prefix=True),
            ],
        )

    def _format_col_range(self, col_start, col_end, col_range):
        """Formats a column-only range."""
        if col_end is None:
            return self._format_single_column(col_start, col_range)
        return self._format_column_span(col_start, col_end, col_range)

    def _format_single_column(self, col_start, col_range):
        """Formats a single column, either numeric or named."""
        if col_range[col_start] is None:
            return self.expand_ref(xl_col_to_name(col_start, col_abs=self.col_start_is_abs))
        return self.expand_ref(col_range[col_start], self.col_start_is_abs)

    def _format_column_span(self, col_start, col_end, col_range):
        """Formats a range of columns."""
        if col_range[col_start] is None:
            return f"{self.expand_ref(xl_col_to_name(col_start, col_abs=self.col_start_is_abs))}:{self.expand_ref(xl_col_to_name(col_end, col_abs=self.col_end_is_abs), no_prefix=True)}"
        return ":".join(
            [
                self.expand_ref(
                    col_range[col_start],
                    self.col_start_is_abs,
                    no_prefix=col_range[col_start].scope == RefScope.DOCUMENT
                    or col_range[col_end].scope == RefScope.DOCUMENT,
                ),
                self.expand_ref(col_range[col_end], self.col_end_is_abs, no_prefix=True),
            ],
        )

    def _format_cell_range(self, row_start, col_start, row_end, col_end):
        """Formats a full cell range."""
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
    def __init__(self, model):
        self.model = model
        self.doc_names: dict[str, int | CellRange] = {}
        self.sheet_names: dict[int, dict[str, int | CellRange]] = {}
        self.table_names: dict[int, dict[str, int | CellRange]] = {}
        self.row_ranges = {}
        self.col_ranges = {}
        self._all_table_names = model.table_names()
        self.unique_tables: dict[str, id] = {}

    @staticmethod
    def _exact_count(pool: list, value: int | str | bool):
        return sum(1 for x in pool if type(x) is type(value) and x == value)

    def _row_data(self, table_id: int, row: int) -> int | str | bool | None:
        num_header_cols = self.model.num_header_cols(table_id)
        return self.model._table_data[table_id][row][num_header_cols - 1].formatted_value

    def _column_data(self, table_id: int, col: int) -> int | str | bool | None:
        num_header_rows = self.model.num_header_rows(table_id)
        return self.model._table_data[table_id][num_header_rows - 1][col].formatted_value

    def _calculate_name_scopes(
        self,
        sheet_id: int,
        table_id: int,
        axis: TableAxis,
    ) -> dict[int, str | None]:
        num_header_rows = self.model.num_header_rows(table_id)
        num_header_cols = self.model.num_header_cols(table_id)

        if axis == TableAxis.ROW:
            data_lookup = self._row_data
            range_start = num_header_rows
            range_end = self.model.number_of_rows(table_id)
            first_offset = num_header_rows
        else:
            data_lookup = self._column_data
            range_start = num_header_cols
            range_end = self.model.number_of_columns(table_id)
            first_offset = num_header_cols

        if (axis == TableAxis.ROW and num_header_cols == 0) or (
            axis == TableAxis.COLUMN and num_header_rows == 0
        ):
            return {idx: None for idx in range(range_end)}

        scopes = {}
        names = []
        all_names = [data_lookup(table_id, idx) for idx in range(range_start, range_end)]
        for idx in range(range_end):
            if idx < first_offset:
                names.append(None)
                scopes[idx] = None
            else:
                name = data_lookup(table_id, idx)
                if name is None:
                    scopes[idx] = None
                elif self._exact_count(all_names, name) > 1:
                    names.append(None)
                    scopes[idx] = None
                else:
                    names.append(name)
                    scopes[idx] = name

        for name in names:
            self.doc_names[name] += 1
            self.sheet_names[sheet_id][name] += 1

        return scopes

    def _calculate_scope_types(
        self,
        sheet_id: int,
        table_id: int,
        axis: TableAxis,
        scopes: dict[int, CellRange | None],
    ):
        """
        For any locally unique row/column names, tag whether they are table-unique,
        sheet-unique or document-unique names.
        """
        if axis == TableAxis.ROW:
            axis_range = range(self.model.number_of_rows(table_id))
        else:
            axis_range = range(self.model.number_of_columns(table_id))

        table_name = self.model.table_name(table_id)
        for idx in axis_range:
            name = scopes[idx]
            scope = ScopedNameRef(name, axis=axis, table_id=table_id, offset=idx)
            if name is None:
                continue
            if name in self.doc_names:
                scopes[idx] = replace(scope, scope=RefScope.DOCUMENT)
                self.doc_names[name] = scopes[idx]
            elif name in self.sheet_names[sheet_id]:
                scopes[idx] = replace(scope, scope=RefScope.SHEET)
                self.sheet_names[sheet_id][name] = scopes[idx]
            else:
                scope_type = (
                    RefScope.TABLE
                    if self._exact_count(self._all_table_names, table_name) == 1
                    else RefScope.NONE
                )
                scopes[idx] = replace(scope, scope=scope_type)
                self.table_names[table_id][name] = scopes[idx]

    def _calculate_table_name_maps(self) -> dict[str, int]:
        self.sheet_name_to_id = {self.model.sheet_name(sid): sid for sid in self.model.sheet_ids()}
        self.sheet_id_to_name = {sid: self.model.sheet_name(sid) for sid in self.model.sheet_ids()}
        table_names = self.model.table_names()
        self.unique_table_name_to_id = {
            self.model.table_name(tid): tid
            for tid in self.model.table_ids()
            if table_names.count(self.model.table_name(tid)) == 1
        }
        self.sheet_table_name_to_id = {
            self.model.sheet_name(sid): {
                self.model.table_name(tid): tid for tid in self.model.table_ids(sid)
            }
            for sid in self.model.sheet_ids()
        }

    def _scoped_ref_to_cell_ref(self, ref: ScopedNameRef) -> CellRange:
        return CellRange(
            model=self.model,
            to_table_id=ref.table_id,
            row_start=ref.offset if ref.axis == TableAxis.ROW else None,
            col_start=ref.offset if ref.axis == TableAxis.COLUMN else None,
        )

    def _deref_doc_scope(self, from_table_id: int, name: str) -> CellRange:
        """Try and use a name reference in the document scope."""
        # Try using the name as document scope
        if name in self.doc_names:
            return self._scoped_ref_to_cell_ref(self.doc_names[name])
        from_sheet_id = self.model.table_id_to_sheet_id(from_table_id)
        if self._name_in_cell_range(name, self.sheet_names[from_sheet_id]):
            return self._scoped_ref_to_cell_ref(self.sheet_names[from_sheet_id][name])
        msg = f"'{name}' does not exist or scope is ambiguous"
        raise ValueError(msg)

    @staticmethod
    def _name_in_cell_range(name: str, cell_range: list[CellRange]) -> bool:
        names = [x.name if isinstance(x, ScopedNameRef) else None for x in cell_range.values()]
        return name in names

    def _deref_single_scope(self, from_table_id: int, name_scope: str, name: str) -> CellRange:
        # Try using the name as sheet/table scope
        from_sheet_id = self.model.table_id_to_sheet_id(from_table_id)
        if name_scope in self.sheet_name_to_id:
            # Test if scope is a sheet name and the reference is valid in that
            # sheet's name scope
            to_sheet_id = self.sheet_name_to_id[name_scope]
            if name in self.sheet_names[to_sheet_id]:
                # Name is valid in a sheet scope
                return self._scoped_ref_to_cell_ref(self.sheet_names[to_sheet_id][name])

        from_sheet_name = self.sheet_id_to_name[from_sheet_id]
        if self._name_in_cell_range(name, self.sheet_names[from_sheet_id]):
            # Scope is a table but in the same sheet so sheet
            # scoping rules apply.
            return self._scoped_ref_to_cell_ref(self.sheet_names[from_sheet_id][name])

        if name_scope in self.unique_table_name_to_id:
            to_table_id = self.unique_table_name_to_id[name_scope]
            if self._name_in_cell_range(name, self.table_names[to_table_id]):
                # Name is valid in table scope and table name is document-unique
                return self._scoped_ref_to_cell_ref(self.table_names[to_table_id][name])
        elif name_scope in self.sheet_table_name_to_id[from_sheet_name]:
            to_table_id = self.sheet_table_name_to_id[from_sheet_name][name_scope]
            # Name is valid in a table in the current sheet
            if self._name_in_cell_range(name, self.row_ranges[to_table_id]):
                return CellRange(
                    model=self.model,
                    to_table_id=to_table_id,
                    row_start=self.row_ranges[to_table_id],
                )
            if self._name_in_cell_range(name, self.col_ranges[to_table_id]):
                return CellRange(
                    model=self.model,
                    to_table_id=to_table_id,
                    col_start=self.col_ranges[to_table_id],
                )

        msg = f"'{name_scope}::{name}' does not exist or scope is ambiguous"
        raise ValueError(msg)

    def _deref_name(
        self,
        from_table_id: int,
        name_scope_1: str,
        name_scope_2: str,
        name: str,
    ) -> CellRange:
        if not name_scope_1 and not name_scope_2:
            return self._deref_doc_scope(from_table_id, name)

        if not name_scope_1:
            return self._deref_single_scope(from_table_id, name_scope_2, name)

        # Full sheet::table::name scope
        try:
            to_table_id = self.sheet_table_name_to_id[name_scope_1][name_scope_2]
            if self._name_in_cell_range(name, self.row_ranges[to_table_id]):
                return CellRange(
                    model=self.model,
                    to_table_id=to_table_id,
                    row_start=self.row_ranges[to_table_id],
                )
            if self._name_in_cell_range(name, self.col_ranges[to_table_id]):
                return CellRange(
                    model=self.model,
                    to_table_id=to_table_id,
                    col_start=self.col_ranges[to_table_id],
                )
        except KeyError:
            # Catch invalid sheet/table names
            pass

        msg = f"'{name_scope_1}::{name_scope_2}::{name}' does not exist or scope is ambiguous"
        raise ValueError(msg)

    def lookup_named_ref(self, from_table_id: int, ref: CellRange) -> tuple[CellRange]:
        print(
            f"\nXREF: {ref.name_scope_1 or '<none>'}::{ref.name_scope_2 or '<none>'}::{ref.row_start}:{ref.row_end}",
        )
        if ref.row_start and ref.row_end:
            # Numbers will use the reduced scope of one part of a range to scope the other
            # so start:en   d in a document scope will resolve if either of the references can
            # be resolved in that scope.
            start_ref, end_ref = None, None
            with suppress(ValueError):
                start_ref = self._deref_name(
                    from_table_id,
                    ref.name_scope_1,
                    ref.name_scope_2,
                    ref.row_start,
                )
            with suppress(ValueError):
                end_ref = self._deref_name(
                    from_table_id,
                    ref.name_scope_1,
                    ref.name_scope_2,
                    ref.row_end,
                )

            if start_ref is None and end_ref is None:
                msg = f"'{ref.name_scope_1}::{ref.name_scope_2}::{ref.row_start}:{ref.row_end}'"
                msg += "does not exist or scope is ambiguous"
                raise ValueError(msg)

            if start_ref is None:
                row_start = [
                    v.offset
                    for k, v in self.row_ranges[end_ref.to_table_id].items()
                    if v is not None and v.name == ref.row_start
                ]
                col_start = [
                    v.offset
                    for k, v in self.col_ranges[end_ref.to_table_id].items()
                    if v is not None and v.name == ref.row_start
                ]
                if len(row_start) == 0 and len(col_start) == 0:
                    msg = f"'{ref.name_scope_1}::{ref.name_scope_2}::{ref.row_start}'"
                    msg += "does not exist or scope is ambiguous"
                    raise ValueError(msg)

                start_ref = CellRange(
                    model=self.model,
                    to_table_id=end_ref.to_table_id,
                    row_start=row_start[0] if row_start else col_start[0],
                )
            elif end_ref is None:
                row_end = [
                    v.offset
                    for k, v in self.row_ranges[start_ref.to_table_id].items()
                    if v is not None and v.name == ref.row_end
                ]
                col_end = [
                    v.offset
                    for k, v in self.col_ranges[start_ref.to_table_id].items()
                    if v is not None and v.name == ref.row_end
                ]
                if len(row_end) == 0 and len(col_end) == 0:
                    msg = f"'{ref.name_scope_1}::{ref.name_scope_2}::{ref.col_start}'"
                    msg += "does not exist or scope is ambiguous"
                    raise ValueError(msg)
                end_ref = CellRange(
                    model=self.model,
                    to_table_id=start_ref.to_table_id,
                    col_start=row_end[0] if row_end else col_end[0],
                )
            return (start_ref, end_ref)
        if ref.row_start:
            return self._deref_name(
                from_table_id,
                ref.name_scope_1,
                ref.name_scope_2,
                ref.row_start,
            )
        return self._deref_name(
            from_table_id,
            ref.name_scope_1,
            ref.name_scope_2,
            ref.col_start,
        )

    def calculate_named_ranges(self):
        """
        Find the globally unique row and column headers and the table unique
        row and column headers for use in range references. Returns a dict
        mapping table ID to lists of rows and columns and their names if
        they are unique.
        """
        self.doc_names = defaultdict(int)
        self.sheet_names = {}
        self._calculate_table_name_maps()

        self.row_ranges = {}
        self.col_ranges = {}
        for sheet_id in self.model.sheet_ids():
            self.sheet_names[sheet_id] = defaultdict(int)
            for table_id in self.model.table_ids(sheet_id):
                self.table_names[table_id] = defaultdict(int)
                self.row_ranges[table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                )
                self.col_ranges[table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                )

        # Re-init the list of document-scoped names if they are unique
        self.doc_names = {name: None for name, count in self.doc_names.items() if count == 1}

        for sheet_id in self.model.sheet_ids():
            # Re-init the list of sheet-scoped names if they are unique
            self.sheet_names[sheet_id] = {
                name: None for name, count in self.sheet_names[sheet_id].items() if count == 1
            }

            for table_id in self.model.table_ids(sheet_id):
                self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                    self.row_ranges[table_id],
                )
                self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                    self.col_ranges[table_id],
                )


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
        msg = f"invalid cell reference {cell_str}"
        raise IndexError(msg)

    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    col = 0
    for expn, char in enumerate(reversed(col_str)):
        col += (ord(char) - ord("A") + 1) * (26**expn)

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
        msg = f"row reference {row} below zero"
        raise IndexError(msg)

    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

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
