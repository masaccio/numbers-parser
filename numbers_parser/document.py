import array
import importlib
import os
import re
import sys
import zipfile

from zipfile import ZipFile
from numbers_parser.codec import IWAFile
from numbers_parser.containers import ItemsList, ObjectStore


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

X_ROW = 0


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
        self._row_cell_counts = [o.cell_count for o in object_store[self._tile_id].rowInfos]
        self._cell_offsets = [
            array.array("h", o.cell_offsets).tolist()
            for o in object_store[self._tile_id].rowInfos
        ]

    @property
    def data(self):
        global X_ROW
        x_cols = [
            [9, 2, 10, 3, 4, 11, 0, 5, 6, 1, 7, 8, 12, 13],
            [10, 12, 3, 4, 14, 0, 5, 6, 7, 1, 8, 11, 12, 2, 9],
        ]
        for row_id in range(len(self._object_store[self._tile_id].rowInfos)):
            cell_storage_buffer = (
                self._object_store[self._tile_id].rowInfos[row_id].cell_storage_buffer
            )
            for col_id in range(int(len(cell_storage_buffer) / 24)):
                hex_str = binascii.hexlify(
                    cell_storage_buffer[col_id * 24 : col_id * 24 + 23], sep=":"
                )
                a = cell_storage_buffer[col_id * 24 + 12]
                b = x_cols[X_ROW].pop(0)
                print(f"({row_id},{col_id} = {hex_str}", a, b, "=", a - b)
            print("\n")
        print("\n")
        X_ROW += 1

        cell_storage_buffers = [
            array.array("b", o.cell_storage_buffer).tolist()
            for o in self._object_store[self._tile_id].rowInfos
        ]

        #  0          1          2            3          4          5          6            7          8            9            10           11           12
        # ['YYY_2_1', 'YYY_3_1', 'YYY_COL_2', 'YYY_1_1', 'YYY_1_2', 'YYY_2_2', 'YYY_ROW_3', 'YYY_3_2', 'YYY_ROW_4', 'YYY_COL_1', 'YYY_ROW_1', 'YYY_ROW_2', 'YYY_4_1', 'YYY_4_2']
        #  .     9        2
        # 10     3        4
        # 11     0        5
        # 6      1        7
        # 8      12       13
        #
        #              YYY_COL_1  YYY_COL_2
        # YYY_ROW_1    YYY_1_1    YYY_1_2
        # YYY_ROW_2    YYY_2_1    YYY_2_2
        # YYY_ROW_3    YYY_3_1    YYY_3_2
        # YYY_ROW_4    YYY_4_1    YYY_4_2

        # 0            1          2         3            4            5         6             7          8          9           10           11          12         13           14
        # ['ZZZ_1_2', 'ZZZ_2_2', 'ZZZ_3_2', 'ZZZ_COL_3', 'ZZZ_ROW_1', 'ZZZ_1_3', 'ZZZ_ROW_2', 'ZZZ_2_1', 'ZZZ_2_3', 'ZZZ_3_3', 'ZZZ_COL_1', 'ZZZ_ROW_3', 'ZZZ_3_1', 'ZZZ_COL_2', 'ZZZ_1_1']
        #
        # .      10     13     3
        # 4      14     0      5
        # 6      7      1      8
        # 11     12     2      9

        # (0,0 = b'05:03:    00:00    :00:00:00:00:08:10:02:00:    0e:00:00:00:05:00:00:00:01:00:00'
        # (0,1 = b'05:03:    73:06    :00:00:00:00:08:10:02:00:    03:00:00:00:05:00:00:00:01:00:00'

        # (1,0 = b'05:03:    00:00    :00:00:00:00:08:10:02:00:    0f:00:00:00:05:00:00:00:01:00:00'
        # (1,1 = b'05:03:    44:04    :00:00:00:00:08:10:02:00:    04:00:00:00:05:00:00:00:01:00:00'
        # (1,2 = b'05:03:    c7:03    :00:00:00:00:08:10:02:00:    05:00:00:00:05:00:00:00:01:00:00'

        # (2,0 = b'05:03:    00:00    :00:00:00:00:08:10:02:00:    10:00:00:00:05:00:00:00:01:00:00'
        # (2,1 = b'05:03:    73:06    :00:00:00:00:08:10:02:00:    01:00:00:00:05:00:00:00:01:00:00'
        # (2,2 = b'05:03:    62:2b    :00:00:00:00:08:10:02:00:    07:00:00:00:05:00:00:00:01:00:00'

        # (3,0 = b'05:03:    00:00    :00:00:00:00:08:10:02:00:    09:00:00:00:05:00:00:00:01:00:00'
        # (3,1 = b'05:03:    40:30    :00:00:00:00:08:10:02:00:    02:00:00:00:05:00:00:00:01:00:00'
        # (3,2 = b'05:03:    35:03    :00:00:00:00:08:10:02:00:    0a:00:00:00:05:00:00:00:01:00:00'

        # (4,0 = b'05:03:    00:00    :00:00:00:00:08:10:02:00:    0c:00:00:00:05:00:00:00:01:00:00'
        # (4,1 = b'05:03:    40:30    :00:00:00:00:08:10:02:00:    11:00:00:00:05:00:00:00:01:00:00'
        # (4,2 = b'05:03:    10:03    :00:00:00:00:08:10:02:00:    12:00:00:00:05:00:00:00:01:00:00'

        data = []
        strings_id = self._table.base_data_store.stringTable.identifier
        table_strings = [x.string for x in self._object_store[strings_id].entries]
        #  TODO: deal with formulas instead of strings
        for row_num in range(self._table.number_of_rows):
            data.append([])
            offsets = self._cell_offsets[row_num]
            print("@" + str(row_num), "offsets =", offsets[:10])
            for col_num in range(self._table.number_of_columns):
                if offsets[col_num] < 0:
                    data[row_num].append("")
                else:
                    data[row_num].append(table_strings.pop(0))
        print([x.string for x in self._object_store[strings_id].entries], "\n")
        return data

    @property
    def num_rows(self):
        return len(self._row_headers)

    @property
    def num_cols(self):
        return len(self._column_headers)


