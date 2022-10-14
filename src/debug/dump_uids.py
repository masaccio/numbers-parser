import argparse

from numbers_parser import Document
from numbers_parser.utils import uuid
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


def find_fields(obj=object, tree=""):
    for descriptor in obj.DESCRIPTOR.fields:
        value = getattr(obj, descriptor.name)
        name = descriptor.name
        if isinstance(value, TSPMessages.UUID) or isinstance(
            value, TSPMessages.CFUUIDArchive
        ):
            if uuid(value) == 0:
                print(f"{tree}.{name} = UNDEFINED")
            else:
                uuid_str = f"{uuid(value):032x}"
                uuid_str = f"0x{uuid_str[0:8]}_{uuid_str[8:16]}_{uuid_str[16:24]}_{uuid_str[24:32]}"
                print(f"{tree}.{name} = {uuid_str}")
        elif (
            not isinstance(value, TSPMessages.Reference)
            and "Archive" in type(value).__module__
        ):
            find_fields(getattr(obj, name), tree + "." + name)


parser = argparse.ArgumentParser()
parser.add_argument("numbers", nargs="*", help="Numbers folders/files to dump")
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    for sheet in doc._sheets:
        for table in sheet._tables:
            print(f"===== sheet={sheet.name} table={table.name}")
            find_fields(doc._model.objects[table._table_id])
