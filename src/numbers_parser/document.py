import array
import importlib
import os
import re
import sys
import struct
import zipfile

from enum import Enum
from datetime import datetime, timedelta

from numbers_parser.containers import ItemsList, ObjectStore, NumbersError
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


import binascii


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
        self._data = None

    @property
    def data(self):
        if self._data is not None:
            return self._data

        row_infos = []
        for tile_id in self._tile_ids:
            row_infos += self._object_store[tile_id].rowInfos

        storage_version = self._object_store[tile_id].storage_version
        if storage_version != 5:
            raise UnsupportedError( # pragma: no cover
                f"Unsupported row info version {storage_version}"
            )

        storage_buffers = [
            extract_cell_data(
                r.cell_storage_buffer, r.cell_offsets, self.num_cols, r.has_wide_offsets
            )
            for r in row_infos
        ]
        storage_buffers_pre_bnc = [
            extract_cell_data(
                r.cell_storage_buffer_pre_bnc,
                r.cell_offsets_pre_bnc,
                self.num_cols,
                r.has_wide_offsets,
            )
            for r in row_infos
        ]

        self._data = []
        num_rows = sum([self._object_store[t].numrows for t in self._tile_ids])
        for row_num in range(num_rows):
            row = []
            for col_num in range(self.num_cols):
                if col_num < len(storage_buffers[row_num]):
                    storage_buffer = storage_buffers[row_num][col_num]
                else:
                    # Rest of row is empty cells
                    row.append(None)
                    continue

                if col_num < len(storage_buffers_pre_bnc[row_num]):
                    storage_buffer_pre_bnc = storage_buffers_pre_bnc[row_num][col_num]
                else:
                    storage_buffer_pre_bnc = None

                if storage_buffer is None:
                    row.append(None)
                else:
                    cell_type = storage_buffer[1]
                    if cell_type == TSTArchives.emptyCellValueType:
                        cell_value = None
                    elif cell_type == TSTArchives.numberCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = None
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                    elif cell_type == TSTArchives.textCellType:
                        key = struct.unpack("<i", storage_buffer[12:16])[0]
                        cell_value = self._table_string(key)
                    elif cell_type == TSTArchives.dateCellType:
                        if storage_buffer_pre_bnc is None:
                            cell_value = None
                        else:
                            seconds = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                            cell_value = datetime(2001, 1, 1) + timedelta(seconds=seconds)
                    elif cell_type == TSTArchives.boolCellType:
                        d = struct.unpack("<d", storage_buffer[12:20])[0]
                        cell_value = d > 0.0
                    elif cell_type == TSTArchives.durationCellType:
                        cell_value = struct.unpack("<d", storage_buffer[12:20])[0]
                    elif cell_type == TSTArchives.formulaErrorCellType:
                        cell_value = "*ERROR*"
                    elif cell_type == 9:
                        cell_value = "*FORMULA*"
                    elif cell_type == 10:
                        if storage_buffer_pre_bnc is None:
                            cell_value = None
                        else:
                            cell_value = struct.unpack("<d", storage_buffer_pre_bnc[-12:-4])[0]
                    else:
                        raise UnsupportedError( # pragma: no cover
                            f"Unsupport cell type {cell_type} @{self.name}:({row_num},{col_num})"
                        )

                    row.append(cell_value)
            self._data.append(row)

        return self._data

    def _table_string(self, key):
        if not hasattr(self, "_table_strings"):
            strings_id = self._table.base_data_store.stringTable.identifier
            self._table_strings = {x.key: x.string for x in self._object_store[strings_id].entries}
        return self._table_strings[key]

    @property
    def num_rows(self):
        return len(self._row_headers)

    @property
    def num_cols(self):
        return len(self._column_headers)

    def cell(self, *args):
        if type(args[0]) == str:
            (row_num, col_num) = xl_cell_to_rowcol(args[0])
        elif len(args) != 2:
            raise IndexError("invalid cell reference " + str(args))
        else:
            (row_num, col_num) = args
        data = self.data
        if row_num >= len(data):
            raise IndexError(f"row {row_num} out of range")
        if col_num >= len(data[row_num]):
            raise IndexError(f"col {col_num} out of range")
        return data[row_num][col_num]

    def iterrows(self, min_row=None, max_row=None):
        data = self.data
        if min_row is None:
            min_row = 0
        if max_row is None:
            max_row = self.num_rows
        for row_num in range(min_row, max_row):
            yield data[row_num]

    def itercols(self, min_col=None, max_col=None):
        data = self.data
        if min_col is None:
            min_col = 0
        if max_col is None:
            max_col = self.num_cols
        for col_num in range(min_col, max_col):
            yield [row[col_num] for row in data]


def extract_cell_data(storage_buffer, offsets, num_cols, has_wide_offsets):
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
            #  https://stackoverflow.com/questions/19502378/
            idx = next((i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None)
            if idx is None:
                end = len(storage_buffer)
            else:
                end = offsets[col_num + idx + 1]
        data.append(storage_buffer[start:end])

    return data


#  Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str):
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
