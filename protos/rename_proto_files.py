import os
import sys
from glob import glob


def rename_proto_files(_dir):
    replacements = {}
    for proto_file in glob(os.path.join(_dir, "*.proto")):
        no_ext = proto_file.replace(".proto", "")
        replaced = no_ext.replace(".", "_") + ".proto"
        replacements[os.path.basename(proto_file)] = os.path.basename(replaced)
        if proto_file != replaced:
            os.rename(proto_file, replaced)

    for proto_file in glob(os.path.join(_dir, "*.proto")):
        original_contents = open(proto_file).read()
        contents = original_contents
        for old, new in replacements.items():
            contents = contents.replace('import "%s";' % old, 'import "%s";' % new)
        if contents != original_contents:
            with open(proto_file, 'w') as f:
                f.write(contents)

rename_proto_files(sys.argv[-1])
