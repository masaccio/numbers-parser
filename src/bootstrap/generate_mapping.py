import json
import os
import sys


if len(sys.argv) != 3:
    raise (ValueError(f"Usage: {sys.argv[0]} mapping.json mapping.py"))

mapping_json = sys.argv[1]
mapping_py = sys.argv[2]

modules = []
module_import_output = ""
module_files_output = ""
mapping_output = ""
for filename in os.listdir("src/numbers_parser/generated"):
    if "pb2" in filename:
        module = filename.replace("_pb2.py", "")
        modules.append(module)
        module_import_output += (
            f"from numbers_parser.generated import {module}_pb2 as {module}\n"
        )
        module_files_output += f"    {module},\n"


with open(mapping_json) as fh:
    mappings = json.load(fh)

for index, symbol in sorted(mappings.items(), key=lambda x: int(x[0])):
    mapping_output += f'    "{index}": "{symbol}",\n'

OUTPUT_CODE = f"""
{module_import_output}

PROTO_FILES = [
{module_files_output}
]

TSPRegistryMapping = {{
{mapping_output}
}}


def compute_maps():
    name_class_map = {{}}

    def add_nested_types(message_type):
        for name in dict(message_type.DESCRIPTOR.nested_types_by_name):
            child_type = getattr(message_type, name)
            name_class_map[child_type.DESCRIPTOR.full_name] = child_type
            add_nested_types(child_type)

    for file in PROTO_FILES:
        for message_name in dict(file.DESCRIPTOR.message_types_by_name):
            message_type = getattr(file, message_name)
            name_class_map[message_type.DESCRIPTOR.full_name] = message_type
            add_nested_types(message_type)

    id_name_map = {{}}
    name_id_map = {{}}
    for k, v in list(TSPRegistryMapping.items()):
        if v in name_class_map: # pragma: no branch
            id_name_map[int(k)] = name_class_map[v]
            if v not in name_id_map:
                name_id_map[v] = int(k)

    return name_class_map, id_name_map, name_id_map


NAME_CLASS_MAP, ID_NAME_MAP, NAME_ID_MAP = compute_maps()
"""

with open(mapping_py, "w") as fh:
    fh.write(OUTPUT_CODE)
