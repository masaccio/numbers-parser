"""Rewrite generated protobuf imports in Python files."""
import os
import re
import sys

for filename in sys.argv[1:]:
    new_filename = filename + ".new"
    old_f = open(filename)
    new_f = open(new_filename, "w")
    new_f.writelines(
        re.sub("^import T", "import numbers_parser.generated.T", line)
        for line in old_f
    )

    old_f.close()
    new_f.close()

    os.rename(new_filename, filename)
