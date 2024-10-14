import os
import sys
from glob import glob


def rename_proto_files(_dir):
    replacements = {}
    for old_path in glob(os.path.join(_dir, "*.proto")):
        old_file_no_ext = os.path.basename(old_path).replace(".proto", "")
        new_file = old_file_no_ext.replace(".", "_") + ".proto"
        replacements[os.path.basename(old_path)] = new_file
        new_path = os.path.join(os.path.dirname(old_path), new_file)
        if old_path != new_path:
            os.rename(old_path, new_path)

    for proto_file in glob(os.path.join(_dir, "*.proto")):
        original_contents = open(proto_file).read()
        contents = original_contents
        for old, new in replacements.items():
            contents = contents.replace(f'import "{old}";', f'import "{new}";')
        if contents != original_contents:
            with open(proto_file, "w") as f:
                f.write(contents)


rename_proto_files(sys.argv[-1])
