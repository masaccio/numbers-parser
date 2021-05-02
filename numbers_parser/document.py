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
            h.numberOfCells
            for h in object_store[bds.rowHeaders.buckets[0].identifier].headers
        ]
        self._column_headers = [
            h.numberOfCells for h in object_store[bds.columnHeaders.identifier].headers
        ]
        self._tile_id = bds.tiles.tiles[0].tile.identifier
        self._cell_offsets = [
            array.array("h", o.cell_offsets).tolist()
            for o in object_store[self._tile_id].rowInfos
        ]
        self._cell_offsets_pre_bnc = [
            array.array("h", o.cell_offsets).tolist()
            for o in object_store[self._tile_id].rowInfos
        ]

    @property
    def data(self):
        strings_id = self._table.base_data_store.stringTable.identifier
        table_strings = {x.key: x.string for x in self._object_store[strings_id].entries}
        data = []
        for row_num in range(self.num_rows):
            row = []
            row_info = self._object_store[self._tile_id].rowInfos[row_num]
            cell_storage_buffers = extract_cell_data(
                row_info.cell_storage_buffer, row_info.cell_offsets, self.num_cols
            )
            cell_storage_buffers_pre_bnc = extract_cell_data(
                row_info.cell_storage_buffer_pre_bnc,
                row_info.cell_offsets_pre_bnc,
                self.num_cols,
            )

            for col_num in range(self.num_cols):
                if cell_storage_buffers[col_num] is None:
                    row.append(None)
                else:
                    if row_info.storage_version != 5:
                        raise UnsupportedError(f"Unsupported row info version {row_info.storage_version}")

                    cell_type = cell_storage_buffers[col_num][1]
                    cell_value = None
                    if cell_type == TSTArchives.numberCellType:
                        cell_value = struct.unpack("<d", cell_storage_buffers_pre_bnc[col_num][24:32])[0]
                    elif cell_type == TSTArchives.textCellType:
                        key = struct.unpack("<i", cell_storage_buffers[col_num][12:16])[0]
                        cell_value = table_strings[key]
                    elif cell_type == TSTArchives.dateCellType:
                        seconds = struct.unpack("<d", cell_storage_buffers_pre_bnc[col_num][24:32])[0]
                        cell_value = datetime(2001,1,1) + timedelta(seconds=seconds)
                    elif cell_type == TSTArchives.boolCellType:
                        d = struct.unpack("<d", cell_storage_buffers[col_num][12:20])[0]
                        cell_value = d > 0.0
                    elif cell_type == TSTArchives.durationCellType:
                        cell_value = struct.unpack("<d", cell_storage_buffers[col_num][12:20])[0]
                    else:
                        raise UnsupportedError(f"Unsupport cell type {cell_type}")

                    row.append(cell_value)
            data.append(row)

        return data

    @property
    def num_rows(self):
        return len(self._row_headers)

    @property
    def num_cols(self):
        return len(self._column_headers)


def extract_cell_data(storage_buffer, offsets, num_cols):
    offsets = array.array("h", offsets).tolist()
    data = []
    for col_num in range(num_cols):
        start = offsets[col_num]
        if start < 0:
            data.append(None)
            continue

        # Get next offset past current one that is not -1
        # https://stackoverflow.com/questions/19502378/
        idx = next((i for i, x in enumerate(offsets[col_num + 1 :]) if x >= 0), None)
        if idx is None:
            end = len(storage_buffer) - 1
        else:
            end = offsets[col_num + idx + 1] - 1
        data.append(storage_buffer[start:end])

    return data
