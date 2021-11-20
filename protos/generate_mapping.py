import os
import sys


if len(sys.argv) != 2:
    print("usage:", sys.argv[0], "lldb-dump", file=sys.stderr)
    sys.exit(1)

modules = []
for filename in os.listdir("src/numbers_parser/generated"):
    if "pb2" in filename:
        module = filename.replace("_pb2.py", "")
        modules.append(module)
        print(f"from numbers_parser.generated import {module}_pb2 as {module}")

print("\n")
print("PROTO_FILES = [")
for module in modules:
    print(f"    {module},")
print("]")
print("\n")

# Expecting LLDB output like this:
#
# (lldb) po [TSPRegistry sharedRegistry]
# <TSPRegistry 0x6000029722f0
#  _messageTypeToPrototypeMap = {
#         7 -> 0x103098480 TN.PlaceholderArchive
#         601 -> 0x1058dcfd8 TSA.FunctionBrowserStateArchive
#         10024 -> 0x108760460 TSWP.DropCapStyleArchive
#         2405 -> 0x108761178 TSWP.StyleReorderCommandArchive
#         2400 -> 0x108761098 TSWP.StyleBaseCommandArchive
#         2104 -> 0x0 null
#         2101 -> 0x108760bc0 TSWP.TextCommandArchive
with open(sys.argv[1]) as fh:
    dump_lines = [line.strip() for line in fh]

in_struct = False
mappings = {}
for line in dump_lines:
    if line.endswith(" = {"):
        in_struct = True
    elif line.endswith("}"):
        break
    elif in_struct:
        args = line.split(" -> ")
        index = args[0]
        symbol = args[1].split()[-1]
        if symbol != "null":
            mappings[int(index)] = symbol


print("TSPRegistryMapping = {")
for index, symbol in sorted(mappings.items()):
    print(f'    "{index}": "{symbol}",')

print(
    """}


def compute_maps():
    name_class_map = {}
    for file in PROTO_FILES:
        for message_name in file.DESCRIPTOR.message_types_by_name:
            message_type = getattr(file, message_name)
            name_class_map[message_type.DESCRIPTOR.full_name] = message_type

    id_name_map = {}
    for k, v in list(TSPRegistryMapping.items()):
        if v in name_class_map:
            id_name_map[int(k)] = name_class_map[v]

    return name_class_map, id_name_map


NAME_CLASS_MAP, ID_NAME_MAP = compute_maps()"""
)
