import argparse

from numbers_parser import Document
from numbers_parser.constants import DOCUMENT_ID
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages
from numbers_parser.numbers_uuid import NumbersUUID

MAX_DEPTH = 10


def deep_print(objects, obj, indent=0, max_depth=MAX_DEPTH):
    if indent >= max_depth:
        return
    for descriptor in obj.DESCRIPTOR.fields:
        value = getattr(obj, descriptor.name)
        if isinstance(value, TSPMessages.Reference):
            obj_id = value.identifier
            name = descriptor.name
            if obj_id != 0:
                type_name = objects[id].DESCRIPTOR.full_name
                print("  " * indent + f"{name}->#{obj_id} [{type_name}]")
                deep_print(objects, objects[id], indent=indent + 1, max_depth=max_depth)
            else:
                print("  " * indent + f"{name}->#{obj_id}")
        elif isinstance(value, (TSPMessages.CFUUIDArchive, TSPMessages.UUID)):
            uuid = NumbersUUID(value)
            print("  " * indent + f"{descriptor.name}={uuid}")
        elif descriptor.type == descriptor.TYPE_MESSAGE:
            if descriptor.label == descriptor.LABEL_REPEATED:
                for i, item in enumerate(value):
                    print("  " * indent + f"{descriptor.name}[{i}]")
                    deep_print(objects, item, indent=indent + 1, max_depth=max_depth)
            else:
                print("  " * indent + f"{descriptor.name}")
                deep_print(objects, value, indent=indent + 1, max_depth=max_depth)
        elif descriptor.type == descriptor.TYPE_ENUM:
            if type(value).__name__ == "RepeatedScalarContainer":
                enum_str = [descriptor.enum_type.values[x].name for x in value]  # noqa: PD011
            else:
                enum_str = descriptor.enum_type.values[value - 1].name  # noqa: PD011
            print("  " * indent + f"{descriptor.name}={enum_str}")
        else:
            print("  " * indent + f"{descriptor.name}={value}")


parser = argparse.ArgumentParser()
parser.add_argument(
    "numbers",
    nargs="*",
    metavar="numbers-filename",
    help="Numbers folders/files to dump",
)
parser.add_argument(
    "--max-depth",
    type=int,
    default=MAX_DEPTH,
    help=f"Maximum reference indirections (default {MAX_DEPTH})",
)
commands = parser.add_mutually_exclusive_group()
commands.add_argument(
    "--document",
    action="store_true",
    help="Dump from document root",
)
commands.add_argument(
    "--tables",
    action="store_true",
    help="Dump all tables",
)
commands.add_argument(
    "--iwa",
    metavar="iwa-file",
    help="Dump all objects in an IWA file",
)
commands.add_argument(
    "--archive",
    metavar="archive-name",
    help="Dump all objects matching the archive name",
)
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    if args.document:
        deep_print(
            doc._model.objects,
            doc._model.objects[DOCUMENT_ID],
            max_depth=args.max_depth,
        )
    elif args.tables:
        for sheet in doc._sheets:
            for table in sheet._tables:
                print(f"===== sheet={sheet.name} table={table.name}")
                deep_print(
                    doc._model.objects,
                    doc._model.objects[table._table_id],
                    max_depth=args.max_depth,
                )
    elif args.iwa:
        filenames = [x for x in doc._model.file_store if args.iwa in x]
        for sub_filename in filenames:
            object_ids = [
                obj_id
                for obj_id in doc._model.objects._object_to_filename_map
                if doc._model.objects._object_to_filename_map[obj_id] == sub_filename
            ]
            for obj_id in sorted(object_ids):
                print(f"===== object={obj_id}")
                deep_print(doc._model.objects, doc._model.objects[obj_id], max_depth=args.max_depth)
    else:
        for obj_id in doc._model.find_refs(args.archive):
            print(f"===== object={obj_id}")
            deep_print(doc._model.objects, doc._model.objects[obj_id], max_depth=args.max_depth)
