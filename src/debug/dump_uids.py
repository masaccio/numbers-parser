import argparse

from numbers_parser import Document
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


def find_fields(obj=object, tree=""):
    for descriptor in obj.DESCRIPTOR.fields:
        value = getattr(obj, descriptor.name)
        name = descriptor.name
        if isinstance(value, TSPMessages.UUID) or isinstance(
            value, TSPMessages.CFUUIDArchive
        ):
            uuid = NumbersUUID(value)
            print(f"{tree}.{name} = {uuid}")
        elif not isinstance(value, TSPMessages.Reference) and hasattr(
            value, "DESCRIPTOR"
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
