import array
import importlib
import os
import re
import sys
import struct
import zipfile

from typing import Union
from typing import Generator
from datetime import datetime, timedelta

from numbers_parser.containers import ItemsList, ObjectStore, NumbersError
from numbers_parser.cell import *
from numbers_parser.generated import TSTArchives_pb2 as TSTArchives


class UnsupportedError(NumbersError):
    """Raised for unsupported file format features"""

    pass


class Document:
    def __init__(self, filename):
        self._object_store = ObjectStore(filename)

    def sheets(self):
        refs = [o.identifier for o in self._object_store[1].sheets]
        self._sheets = Sheets(self._object_store, refs)
        return self._sheets


class Sheets(ItemsList):
    def __init__(self, object_store, refs):
        super(Sheets, self).__init__(object_store, refs, Sheet)


class Sheet:
    def __init__(self, object_store, sheet_id):
        self._sheet_id = sheet_id
        self._object_store = object_store
        self._sheet = object_store[sheet_id]
        self.name = self._sheet.name

    def tables(self):
        table_ids = self._object_store.find_refs("TableInfoArchive")
        table_refs = [
            self._object_store[table_id].tableModel.identifier
            for table_id in table_ids
            if self._object_store[table_id].super.parent.identifier == self._sheet_id
        ]
        self._tables = Tables(self._object_store, table_refs)
        return self._tables


class Tables(ItemsList):
    def __init__(self, object_store, refs):
        super(Tables, self).__init__(object_store, refs, Table)


class Table:
    def __init__(self, object_store, table_id):
        super(Table, self).__init__()
        self._object_store = object_store
        self._table = object_store[table_id]
        self._table_id = table_id
        self.name = self._table.table_name

        bds = self._table.base_data_store
        self._row_headers = [
            h.numberOfCells for h in object_store[bds.rowHeaders.buckets[0].identifier].headers
        ]
        self._column_headers = [
            h.numberOfCells for h in object_store[bds.columnHeaders.identifier].headers
        ]
        self._tile_ids = [t.tile.identifier for t in bds.tiles.tiles]
        self._table_data = None
        self._table_cells = None

    @property
    def data(self):
        if self._table_data is None:
            table_data = []
            for row in self.cells:
                table_data.append([cell.value for cell in row])
            self._table_data = table_data

        return self._table_data

    @property
    def cells(self) -> list:
        if self._table_cells is None:
            self._table_cells = self._extract_table_cells()
        return self._table_cells

    def _extract_table_cells(self):
        row_infos = []
        for tile_id in self._tile_ids:
            row_infos += self._object_store[tile_id].rowInfos

        storage_version = self._object_store[tile_id].storage_version
        if storage_version != 5:
            raise UnsupportedError(  # pragma: no cover
                f"Unsupported row info version {storage_version}"
            )

        storage_buffers = [
            _extract_cell_data(
                r.cell_storage_buffer, r.cell_offsets, self.num_cols, r.has_wide_offsets
            )
            for r in row_infos
        ]
        storage_buffers_pre_bnc = [
            _extract_cell_data(
                r.cell_storage_buffer_pre_bnc,
                r.cell_offsets_pre_bnc,
                self.num_cols,
                r.has_wide_offsets,
            )
            for r in row_infos
        ]

        table_cells = []
        num_rows = sum([self._object_store[t].numrows for t in self._tile_ids])
        for row_num in range(num_rows):
            row = []
            for col_num in range(self.num_cols):
                if col_num < len(storage_buffers[row_num]):
                    storage_buffer = storage_buffers[row_num][col_num]
                else:
                    # Rest of row is empty cells
                    row.append(EmptyCell(row_num, col_num))
                    continue

                if col_num < len(storage_buffers_pre_bnc[row_num]):
                    storage_buffer_pre_bnc = storage_buffers_pre_bnc[row_num][col_num]
                else:
                    storage_buffer_pre_bnc = None

                if storage_buffer is None:
                    row.append(
                        MergedCell(
                            # TODO: include full range information
                            row_start=row_num,
                            row_end=row_num,
                            col_start=col_num,
                            col_end=col_num,
                        )
                    )
                else:
                    cell_type = storage_buffer[1]
                    if cell_type == TSTArchives.emptyCellValueType:
                        row.append(EmptyCell(row_num, col_num))
                    elif cell_type == TSTArchives.numberCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                        row.append(NumberCell(row_num, col_num, cell_value))
                    elif cell_type == TSTArchives.textCellType:
                        key = struct.unpack("<i", storage_buffer[12:16])[0]
                        row.append(TextCell(row_num, col_num, self._table_string(key)))
                    elif cell_type == TSTArchives.dateCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = datetime(2001, 1, 1)
                        else:
                            seconds = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                            cell_value = datetime(2001, 1, 1) + timedelta(seconds=seconds)
                        row.append(DateCell(row_num, col_num, cell_value))
                    elif cell_type == TSTArchives.boolCellType:
                        d = struct.unpack("<d", storage_buffer[12:20])[0]
                        row.append(BoolCell(row_num, col_num, d > 0.0))
                    elif cell_type == TSTArchives.durationCellType:
                        cell_value = struct.unpack("<d", storage_buffer[12:20])[0]
                        row.append(DurationCell(row_num, col_num, timedelta(days=cell_value)))
                    elif cell_type == TSTArchives.formulaErrorCellType:
                        row.append(ErrorCell(row_num, col_num))
                    elif cell_type == 9:
                        row.append(FormulaCell(row_num, col_num))
                    elif cell_type == 10:
                        # TODO: Is this realy a number cell (experiments suggest so)
                        if storage_buffer_pre_bnc is None:
                            cell_value = 0.0
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                        row.append(NumberCell(row_num, col_num, cell_value))
                    else:
                        raise UnsupportedError(  # pragma: no cover
                            f"Unsupport cell type {cell_type} @{self.name}:({row_num},{col_num})"
                        )
            table_cells.append(row)

        return table_cells

    def _table_string(self, key):
        if not hasattr(self, "_table_strings"):
            strings_id = self._table.base_data_store.stringTable.identifier
            self._table_strings = {x.key: x.string for x in self._object_store[strings_id].entries}
        return self._table_strings[key]

    @property
    def num_rows(self) -> int:
        return len(self._row_headers)

    @property
    def num_cols(self) -> int:
        return len(self._column_headers)

    def cell(self, *args) -> Union[Cell, MergedCell]:
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
        elif len(args) != 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args

        cells = self.cells
        if row_num >= len(cells):
            raise IndexError(f"row {row_num} out of range")
        if col_num >= len(cells[row_num]):
            raise IndexError(f"coumn {col_num} out of range")
        return cells[row_num][col_num]

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
        cells = self.cells
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

        for row_num in range(min_row, max_row + 1):
            if values_only:
                yield tuple(cell.value or None for cell in cells[row_num][min_col : max_col + 1])
            else:
                yield tuple(cells[row_num][min_col : max_col + 1])

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
        cells = self.cells
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

        for col_num in range(min_col, max_col + 1):
            if values_only:
                yield tuple(row[col_num].value for row in cells[min_row : max_row + 1])
            else:
                yield tuple(row[col_num] for row in cells[min_row : max_row + 1])


def _extract_cell_data(
    storage_buffer: bytes, offsets: list, num_cols: int, has_wide_offsets: bool
) -> list[bytes]:
    """
    Extract storage buffers for each cell in a table row
    Args:
        storage_buffer:  cell_storage_buffer or cell_storage_buffer for a table row
        offsets: 16-bit cell offsets for a table row
        num_cols: number of columns in a table row
        has_wide_offsets: use 4-byte offsets rather than 1-byte offset
    Returns:
         data: list of bytes for each cell in a row, or None if empty
    """
    offsets = array.array("h", offsets).tolist()
    if has_wide_offsets:
        offsets = [o * 4 for o in offsets]

    data = []
    for col_num in range(num_cols):
        if col_num >= len(offsets):
            break

        start = offsets[col_num]
        if start < 0:
            data.append(None)
            continue

        if col_num == (len(offsets) - 1):
            end = len(storage_buffer)
        else:
            # Get next offset past current one that is not -1
            # https://stackoverflow.com/questions/19502378/
            idx = next((i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None)
            if idx is None:
                end = len(storage_buffer)
            else:
                end = offsets[col_num + idx + 1]
        data.append(storage_buffer[start:end])

    return data


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
        col += (ord(char) - ord("A") + 1) * (26 ** expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col
