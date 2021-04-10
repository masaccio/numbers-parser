import array
import importlib
import os
import re
import sys
import zipfile

from zipfile import ZipFile
from numbers_parser.codec import IWAFile
from numbers_parser.items_list import ItemsList


class NumbersError(Exception):
    """Base class for other exceptions"""

    pass


class FileError(NumbersError):
    """Raised for IO and other OS errors"""

    pass


class FileFormatError(NumbersError):
    """Raised for parsing errors during file load"""

    pass


class Document:
    def __init__(self, filename):
        self._objects = get_objects_from_file(filename)

    def sheets(self):
        refs = [o.identifier for o in self._objects[1].sheets]
        self._sheets = Sheets(self._objects, refs)
        return self._sheets


class Sheets(ItemsList):
    def __init__(self, objects, refs):
        super(Sheets, self).__init__(objects, refs, Sheet)


class Sheet:
    def __init__(self, objects, sheet_id):
        self._objects = objects
        self._sheet_id = sheet_id
        self._sheet = objects[sheet_id]
        self.name = self._sheet.name

    def tables(self):
        table_refs = find_tables(self._objects, self._sheet_id)
        self._tables = Tables(self._objects, table_refs)
        return self._tables


import binascii

X_ROW = 0


class Tables(ItemsList):
    def __init__(self, objects, refs):
        super(Tables, self).__init__(objects, refs, Table)


class Table:
    def __init__(self, objects, table_id):
        self._objects = objects
        self._table = objects[table_id]
        self._table_id = table_id
        self.name = self._table.table_name

    @property
    def data(self):
        # styles = self._table.base_data_store.styleTable.identifier
        # row_headers = self._table.base_data_store.rowHeaders.buckets[0].identifier
        # column_headers = self._table.base_data_store.columnHeaders.identifier
        # cell_counts = [o.cell_count for o in self._objects[tile_id].rowInfos]

        global X_ROW
        tile_id = self._table.base_data_store.tiles.tiles[0].tile.identifier
        cell_offsets = [o.cell_offsets for o in self._objects[tile_id].rowInfos]
        x_cols = [
            [9, 2, 10, 3, 4, 11, 0, 5, 6, 1, 7, 8, 12, 13],
            [10, 12, 3, 4, 14, 0, 5, 6, 7, 1, 8, 11, 12, 2, 9],
        ]
        for row_id in range(len(self._objects[tile_id].rowInfos)):
            cell_storage_buffer = (
                self._objects[tile_id].rowInfos[row_id].cell_storage_buffer
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
            for o in self._objects[tile_id].rowInfos
        ]

        #  0          1          2            3          4          5          6            7          8            9            10           11           12
        # ['YYY_2_1', 'YYY_3_1', 'YYY_COL_2', 'YYY_1_1', 'YYY_1_2', 'YYY_2_2', 'YYY_ROW_3', 'YYY_3_2', 'YYY_ROW_4', 'YYY_COL_1', 'YYY_ROW_1', 'YYY_ROW_2', 'YYY_4_1', 'YYY_4_2']
        #  .      9        2
        # 10     3        4
        #  11     0        5
        # 6      1        7
        # 8      12       13

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
        table_strings = [x.string for x in self._objects[strings_id].entries]
        #  TODO: deal with formulas instead of strings
        for row_num in range(self._table.number_of_rows):
            data.append([])
            offsets = array.array("h", cell_offsets[row_num]).tolist()
            for col_num in range(self._table.number_of_columns):
                if offsets[col_num] < 0:
                    data[row_num].append("")
                else:
                    data[row_num].append(table_strings.pop(0))
        print([x.string for x in self._objects[strings_id].entries], "\n")
        return data

    @property
    def num_rows(self):
        row_header_id = self._table.base_data_store.rowHeaders.buckets[0].identifier
        return len(self._objects[row_header_id].headers)

    @property
    def num_cols(self):
        column_header_id = self._table.base_data_store.columnHeaders.identifier
        return len(self._objects[column_header_id].headers)


def find_refs(objects, ref_name):
    refs = [k for k, v in objects.items() if type(v).__name__ == ref_name]
    return refs


def find_objects(objects, ref_name, class_name):
    refs = find_refs(objects, ref_name)
    class_ = getattr(importlib.import_module(__name__), class_name)
    return [class_(objects, obj_id) for obj_id in refs]


def find_tables(objects, parent_sheet_id):
    table_ids = find_refs(objects, "TableInfoArchive")
    table_refs = [
        objects[table_id].tableModel.identifier
        for table_id in table_ids
        if objects[table_id].super.parent.identifier == parent_sheet_id
    ]
    return table_refs


def get_objects_from_file(filename):
    try:
        zipf = zipfile.ZipFile(filename)
    except Error as e:
        raise FileError(f"{filename}: " + str(e))

    objects = {}
    iwa_files = filter(lambda x: x.endswith(".iwa"), zipf.namelist())
    for iwa_filename in iwa_files:
        contents = zipf.read(iwa_filename)

        try:
            iwaf = IWAFile.from_buffer(contents, iwa_filename)
        except Error as e:
            raise FileFormatError(f"{filename}: invalid IWA file {iwa_filename}") from e

        if len(iwaf.chunks) != 1:
            raise FileFormatError(f"{filename}: chunk count != 1 in {iwa_filename}")
        for archive in iwaf.chunks[0].archives:
            if len(archive.objects) == 0:
                raise FileFormatError(f"{filename}: no objects in {iwa_filename}")

            identifier = archive.header.identifier
            if identifier in objects:
                raise FileFormatError(f"{filename}: duplicate reference {identifier}")

            if len(archive.objects) == 1:
                objects[identifier] = archive.objects[0]
            else:
                print(f"{iwa_filename}: found", len(archive.objects), "objects")

    return objects
