import argparse
import re

from numbers_parser import Document
from numbers_parser.constants import PACKAGE_ID


def dump_metadata(objects):
    component_map = {c.identifier: c for c in objects[PACKAGE_ID].components}
    for c in sorted(component_map.values(), key=lambda x: x.preferred_locator):
        if len(c.external_references) == 0:
            continue
        print(f"id={c.identifier}, file={c.preferred_locator}")
        for ref in sorted(
            sorted(c.external_references, key=lambda x: x.component_identifier),
            key=lambda x: x.object_identifier,
        ):
            c_type = type(objects[ref.component_identifier])
            c_name = re.sub(r"Archive\w*_pb2", "", c_type.__module__)
            c_name += "." + c_type.__name__.replace("Archive", "")
            c_name = f"component={ref.component_identifier}({c_name})"

            if ref.object_identifier:
                o_type = type(objects[ref.object_identifier])
                o_name = re.sub(r"Archive\w*_pb2", "", o_type.__module__)
                o_name += "." + o_type.__name__.replace("Archive", "")
                o_name = f"object={ref.object_identifier}({o_name})"
                print("  " + o_name + " " + c_name + " " + f"weak={ref.is_weak}")
            else:
                print(c_name)


parser = argparse.ArgumentParser()
parser.add_argument("numbers", nargs="*", help="Numbers folders/files to dump")
args = parser.parse_args()

for filename in args.numbers:
    doc = Document(filename)
    dump_metadata(doc._model.objects)
