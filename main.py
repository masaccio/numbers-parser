import array
import os
import re
import sys
import zipfile

from zipfile import ZipFile
from numbers_parser.codec import IWAFile


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
        self._sheets = []
        for obj in self._objects[1].sheets:
            self._sheets.append(Sheet(self._objects, obj.identifier))
        return self._sheets

    def tables(self):
        self._tables = find_tables(self._objects)
        return self._tables


class Sheet:
    def __init__(self, objects, sheet_id):
        self._objects = objects
        self._sheet = objects[sheet_id]
        self.name = self._sheet.name
        # TODO: which tables are in which sheets?


class Table:
    def __init__(self, objects, table_id):
        self._objects = objects
        self._table = objects[table_id]
        self._table_id = table_id
        self.name = self._table.table_name
        # self._styles = self._table.base_data_store.styleTable.identifier
        # self._row_headers = self._table.base_data_store.rowHeaders.buckets[0].identifier
        # self._column_headers = self._table.base_data_store.columnHeaders.identifier
        
        tile_id = self._table.base_data_store.tiles.tiles[0].tile.identifier
        cell_offsets = [o.cell_offsets for o in self._objects[tile_id].rowInfos]
        cell_counts = [o.cell_count for o in self._objects[tile_id].rowInfos]
        self.table_values = []
        strings_id = self._table.base_data_store.stringTable.identifier
        table_strings = [x.string for x in objects[strings_id].entries]
        # TODO: deal with formulas instead of strings
        for row_num in range(self._table.number_of_rows):
            self.table_values.append([])
            offsets = array.array('h', cell_offsets[row_num]).tolist()
            for col_num in range(self._table.number_of_columns):
                if offsets[col_num] < 0:
                    self.table_values[row_num].append("")
                else:
                    self.table_values[row_num].append(table_strings.pop(0))

        # print(self.name, "- cell_offsets_pre_bnc")
        # print_offset_table(self._objects[tile_id].rowInfos, self._table.number_of_columns, "cell_offsets_pre_bnc")
        # print(self.name, "- cell_offsets")
        # print_offset_table(self._objects[tile_id].rowInfos, self._table.number_of_columns, "cell_offsets")
        # for row in self.table_values:
        #     for col in row:
        #         print("|" + col.ljust(15), end='')
        #     print("")
        # print("")
        patches = {}
        find_references(patches, self._objects, self._table, verbose=True)
        # for ref_id, patch in patches.items():
        #     print(f"= patching {ref_id}", len(table_dump))
        #     matches = re.search(f"^(\s+identifier: {ref_id})$", table_dump, flags=re.MULTILINE)
        #     if matches is not None:
        #         for match in matches.groups():
        #             indent = re.sub("identifier.*", "", match)
        #             patch_dump = str(patches[ref_id])
        #             patch_dump = re.sub("^", indent, patch_dump, flags=re.MULTILINE)
        #             table_dump = re.sub(f"^\s+identifier: {ref_id}", patch_dump, table_dump, flags=re.MULTILINE)

        if ".patches" not in self._objects:
            self._objects[".patches"] = patches
        else:
            self._objects[".patches"].update(patches)


def print_offset_table(t, width, field):
    if "bnc" in field:
        incr = 20
    else:
        incr = 24
    for row in t:
        data = b''
        data = eval(f"row.{field}")
        offsets = array.array("h", data).tolist()
        offsets = offsets[0:width]
        offset = 0
        for col_num in range(len(offsets)):
            if offsets[col_num] >= 0:
                offsets[col_num] = offset - offsets[col_num]
                offset += incr
            elif offsets[col_num] < 0:
                offsets[col_num] = ""
        print(offsets)


def find_tables(objects):
    table_refs = [
        k for k, v in objects.items() if type(v).__name__ == "TableModelArchive"
    ]
    tables = [Table(objects, table_id) for table_id in table_refs]
    return tables


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


def find_references(patches, objects, obj, parent_obj=None, field="", verbose=False):
    obj_type = type(obj).__name__
    obj_module = type(obj).__module__
    if obj_module == "builtins" or obj_type.startswith("builtin"):
        return

    fields = dir(obj)
    if "DESCRIPTOR" in fields:
        fields = obj.DESCRIPTOR.fields_by_name.keys()
    else:
        return

    if obj_type == "Reference" and obj.identifier != 0:
        patches[obj.identifier] = {
                "parent": parent_obj,
                "field": field,
                "identifier": obj.identifier,
            }

    for field in fields:
        expr = f'find_references(patches, objects, obj.{field}, obj, "{field}")'
        try:
            eval(expr)
        except Exception as e:
            print(f"{expr}: eval failed:", str(e))
            sys.exit(1)


if len(sys.argv) > 1:
    filename = sys.argv[1]
else:
    filename = "tests/data/test-spreadsheet.numbers"

print(f"Opening {filename}")
doc = Document(filename)
for sheet in doc.sheets():
    print(sheet.name)

for table in doc.tables():
    print("====", table.name, "\n", str(table._table))
for ref_id, obj in doc._objects.items():
    print("==", ref_id, "\n", str(obj))

print("Done")
