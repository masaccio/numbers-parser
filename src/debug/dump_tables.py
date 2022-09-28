import argparse

from numbers_parser import Document
from numbers_parser.constants import DOCUMENT_ID
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


def deep_print(objects, obj, indent=0):
    for descriptor in obj.DESCRIPTOR.fields:
        value = getattr(obj, descriptor.name)
        if isinstance(value, TSPMessages.Reference):
            id = value.identifier
            name = descriptor.full_name
            if id != 0:
                type_name = objects[id].DESCRIPTOR.full_name
                print("  " * indent + f"{name} -> #{id} ({type_name})")
                deep_print(objects, objects[id], indent=indent + 1)
            else:
                print("  " * indent + f"{name} -> #{id}")
        elif descriptor.type == descriptor.TYPE_MESSAGE:
            if descriptor.label == descriptor.LABEL_REPEATED:
                map(lambda x: deep_print(objects, x, indent + 1), value)
            else:
                print("  " * indent + f"{descriptor.full_name}")
                deep_print(objects, value, indent=indent + 1)
        elif descriptor.type == descriptor.TYPE_ENUM:
            if type(value).__name__ == "RepeatedScalarContainer":
                enum_str = [descriptor.enum_type.values[x].name for x in value]
            else:
                enum_str = descriptor.enum_type.values[value - 1].name
            print("  " * indent + f"{descriptor.full_name}, {enum_str}")
        else:
            print("  " * indent + f"{descriptor.full_name} {value}")


parser = argparse.ArgumentParser()
parser.add_argument("numbers", nargs="*", help="Numbers folders/files to dump")
parser.add_argument(
    "--document", action="store_true", help="Dump from Document root instead of tables"
)
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    if args.document:
        deep_print(doc._model.objects, doc._model.objects[DOCUMENT_ID])
    else:
        for sheet in doc._sheets:
            for table in sheet._tables:
                print(f"===== sheet={sheet.name} table={table.name}")
                deep_print(doc._model.objects, doc._model.objects[table._table_id])
