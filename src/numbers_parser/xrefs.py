from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from enum import IntEnum
from itertools import chain

from numbers_parser.constants import OPERATOR_PRECEDENCE


class TableAxis(IntEnum):
    ROW = 1
    COLUMN = 2


@dataclass
class CellRefType:
    name: str
    is_document_unique: bool = False
    is_table_unique: bool = False
    is_sheet_unique: bool = False


@dataclass
class CellRef:
    model: object = None
    start: tuple[int] = (None, None)
    end: tuple[int] = (None, None)
    start_abs: tuple[bool] = (False, False)
    end_abs: tuple[bool] = (False, False)
    table_ids: tuple[int] = (None, None)

    def __post_init__(self):
        self._initialize_table_data()
        self._set_sheet_ids()
        self.table_ref_engine = TableRefs(self.model)
        self.table_ref_engine.calculate_named_ranges()
        self.row_ranges = self.table_ref_engine.row_ranges
        self.col_ranges = self.table_ref_engine.col_ranges

    def _initialize_table_data(self):
        self.table_names = list(
            chain.from_iterable(
                [self.model.table_name(tid) for tid in self.model.table_ids(sid)]
                for sid in self.model.sheet_ids()
            ),
        )
        self.table_name_unique = {
            name: self.table_names.count(name) == 1 for name in self.table_names
        }

    def _set_sheet_ids(self):
        """Determine the sheet IDs for the referenced tables."""
        if self.table_ids[1] is None:
            self.table_ids = (self.table_ids[0], self.table_ids[0])
        self.sheet_ids = (
            self.model.table_id_to_sheet_id(self.table_ids[0]),
            self.model.table_id_to_sheet_id(self.table_ids[1]),
        )

    def expand_ref(self, ref: str, is_abs: bool = False, no_prefix=False) -> str:
        is_document_unique = ref.is_document_unique if isinstance(ref, CellRefType) else False
        is_sheet_unique = (
            ref.is_sheet_unique and not is_abs if isinstance(ref, CellRefType) else False
        )

        if isinstance(ref, CellRefType):
            ref = f"${ref.name}" if is_abs else ref.name
        else:
            ref = f"${ref}" if is_abs else ref
        if any(x in ref for x in OPERATOR_PRECEDENCE):
            ref = f"'{ref}'"
        elif "'" in ref:
            ref = ref.replace("'", "'''")

        if no_prefix or is_document_unique:
            return ref

        table_name = self.model.table_name(self.table_ids[1])
        sheet_name = self.model.sheet_name(self.sheet_ids[1])
        if self.table_ids[0] != self.table_ids[1]:
            if self.sheet_ids[0] == self.sheet_ids[1] and is_sheet_unique:
                return ref
            ref = f"{table_name}::{ref}"

        is_table_name_unique = self.table_name_unique[table_name]
        if self.sheet_ids[0] != self.sheet_ids[1] and not is_table_name_unique:
            sheet_name = self.model.sheet_name(self.sheet_ids[1])
            ref = f"{sheet_name}::{ref}"

        return ref

    def __str__(self):
        row_start, col_start = self.start
        row_end, col_end = self.end

        # Handle row-only ranges
        if col_start is None:
            row_range = self.row_ranges[self.sheet_ids[1]][self.table_ids[1]]
            return self._format_row_range(row_start, row_end, row_range)

        # Handle column-only ranges
        if row_start is None:
            col_range = self.col_ranges[self.sheet_ids[1]][self.table_ids[1]]
            return self._format_col_range(col_start, col_end, col_range)

        # Handle full cell ranges
        return self._format_cell_range(row_start, col_start, row_end, col_end)

    def _format_row_range(self, row_start, row_end, row_range):
        """Formats a row-only range."""
        if row_end is None:
            return self._format_single_row(row_start, row_range)
        return self._format_row_span(row_start, row_end, row_range)

    def _format_single_row(self, row_start, row_range):
        """Formats a single row, either numeric or named."""
        if row_range[row_start] is None:
            return self._format_numeric_row(row_start)
        return self.expand_ref(row_range[row_start], self.start_abs[0])

    def _format_numeric_row(self, row_start):
        """Formats a single numeric row."""
        return ":".join(
            [
                self.expand_ref(str(row_start + 1), self.start_abs[0]),
                self.expand_ref(str(row_start + 1), self.start_abs[0], no_prefix=True),
            ],
        )

    def _format_row_span(self, row_start, row_end, row_range):
        """Formats a range of rows."""
        if row_range[row_start] is None:
            return ":".join(
                [
                    self.expand_ref(str(row_start + 1), self.start_abs[0]),
                    self.expand_ref(str(row_end + 1), self.end_abs[0], no_prefix=True),
                ],
            )
        return ":".join(
            [
                self.expand_ref(
                    row_range[row_start],
                    self.start_abs[0],
                    no_prefix=row_range[row_start].is_document_unique
                    or row_range[row_end].is_document_unique,
                ),
                self.expand_ref(row_range[row_end], self.end_abs[0], no_prefix=True),
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
            return self.expand_ref(xl_col_to_name(col_start, col_abs=self.start_abs[1]))
        return self.expand_ref(col_range[col_start], self.start_abs[1])

    def _format_column_span(self, col_start, col_end, col_range):
        """Formats a range of columns."""
        if col_range[col_start] is None:
            return f"{self.expand_ref(xl_col_to_name(col_start, col_abs=self.start_abs[1]))}:{self.expand_ref(xl_col_to_name(col_end, col_abs=self.end_abs[1]), no_prefix=True)}"
        return ":".join(
            [
                self.expand_ref(
                    col_range[col_start],
                    self.start_abs[1],
                    no_prefix=col_range[col_start].is_document_unique
                    or col_range[col_end].is_document_unique,
                ),
                self.expand_ref(col_range[col_end], self.end_abs[1], no_prefix=True),
            ],
        )

    def _format_cell_range(self, row_start, col_start, row_end, col_end):
        """Formats a full cell range."""
        if row_end is None or col_end is None:
            return self.expand_ref(
                xl_rowcol_to_cell(
                    row_start,
                    col_start,
                    row_abs=self.start_abs[0],
                    col_abs=self.start_abs[1],
                ),
            )
        return ":".join(
            [
                self.expand_ref(
                    xl_rowcol_to_cell(
                        row_start,
                        col_start,
                        row_abs=self.start_abs[0],
                        col_abs=self.start_abs[1],
                    ),
                ),
                self.expand_ref(
                    xl_rowcol_to_cell(
                        row_end,
                        col_end,
                        row_abs=self.end_abs[0],
                        col_abs=self.end_abs[1],
                    ),
                    no_prefix=True,
                ),
            ],
        )


class TableRefs:
    def __init__(self, model):
        self.model = model

    @property
    def row_ranges(self):
        return self._row_offset_to_name

    @property
    def col_ranges(self):
        return self._col_offset_to_name

    @staticmethod
    def _exact_count(pool: list, value: int | str | bool):
        return sum(1 for x in pool if type(x) is type(value) and x == value)

    def _row_range(self, table_id: int) -> range:
        num_header_rows = self.model.num_header_rows(table_id)
        return range(num_header_rows, len(self.model._table_data[table_id]))

    def _column_range(self, table_id: int) -> range:
        num_header_cols = self.model.num_header_cols(table_id)
        return range(num_header_cols, len(self.model._table_data[table_id][0]))

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

        # for name in names:
        #     if self._exact_count(names, name) > 1:
        #         names = list(filter(lambda x: x != name, names))

        return scopes

    def _calculate_scope_types(
        self,
        sheet_id: int,
        table_id: int,
        axis: TableAxis,
        scopes: dict[int, CellRef | None],
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
            if name is None:
                continue
            if name in self.doc_names:
                scopes[idx] = CellRefType(name, is_document_unique=True)
            elif name in self.sheet_names[sheet_id] and self.sheet_names[sheet_id][name] == 1:
                scopes[idx] = CellRefType(name, is_sheet_unique=True)
            elif self.table_names.count(table_name) > 1:
                scopes[idx] = CellRefType(name)
            else:
                scopes[idx] = CellRefType(name, is_table_unique=True)

    def calculate_named_ranges(self):
        """
        Find the globally unique row and column headers and the table unique
        row and column headers for use in range references. Returns a dict
        mapping table ID to lists of rows and columns and their names if
        they are unique.
        """
        self.doc_names = defaultdict(int)
        self.sheet_names = {}

        row_offset_to_name = {}
        col_offset_to_name = {}
        for sheet_id in self.model.sheet_ids():
            self.sheet_names[sheet_id] = defaultdict(int)
            row_offset_to_name[sheet_id] = {}
            col_offset_to_name[sheet_id] = {}
            for table_id in self.model.table_ids(sheet_id):
                row_offset_to_name[sheet_id][table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                )
                col_offset_to_name[sheet_id][table_id] = self._calculate_name_scopes(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                )

        self.doc_names = [name for name, count in self.doc_names.items() if count == 1]
        self.table_names = list(
            chain.from_iterable(
                [
                    [self.model.table_name(tid) for tid in self.model.table_ids(sid)]
                    for sid in self.model.sheet_ids()
                ],
            ),
        )

        for sheet_id in self.model.sheet_ids():
            for table_id in self.model.table_ids(sheet_id):
                self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.ROW,
                    row_offset_to_name[sheet_id][table_id],
                )
                self._calculate_scope_types(
                    sheet_id,
                    table_id,
                    TableAxis.COLUMN,
                    col_offset_to_name[sheet_id][table_id],
                )

        self._row_offset_to_name = row_offset_to_name
        self._col_offset_to_name = col_offset_to_name


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
