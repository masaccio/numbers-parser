import json
import os
import sys

if len(sys.argv) != 3:
    msg = f"Usage: {sys.argv[0]} mapping.json mapping.py"
    raise (ValueError(msg))

mapping_json = sys.argv[1]
mapping_py = sys.argv[2]

module_import_output = ""
for name in sorted(
    d.name for d in os.scandir("src/numbers_parser/generated") if d.is_dir() and d.name.startswith("T")
):
    module_import_output += f"    {name},\n"
mapping_output = ""

with open(mapping_json) as fh:
    mappings = json.load(fh)

for index, symbol in sorted(mappings.items(), key=lambda x: int(x[0])):
    mapping_output += f'    "{index}": "{symbol}",\n'

OUTPUT_CODE = f"""from numbers_parser.generated import (  # noqa: F401
{module_import_output.rstrip()}
)
from numbers_parser.generated.message_pool import default_message_pool

TSPRegistryMapping = {{
    {mapping_output.strip()}
}}


def compute_maps():
    name_class_map = {{
        url.rsplit("/", 1)[-1]: klass for url, klass in default_message_pool.url_to_type.items()
    }}

    id_name_map = {{}}
    name_id_map = {{}}
    for k, v in list(TSPRegistryMapping.items()):
        if v in name_class_map:  # pragma: no branch
            id_name_map[int(k)] = name_class_map[v]
            if v not in name_id_map:
                name_id_map[v] = int(k)

    return name_class_map, id_name_map, name_id_map


NAME_CLASS_MAP, ID_NAME_MAP, NAME_ID_MAP = compute_maps()
"""

with open(mapping_py, "w") as fh:
    fh.write(OUTPUT_CODE)
