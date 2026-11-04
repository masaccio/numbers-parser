"""Rewrite generated protobuf imports in Python files."""
import os
import re
import sys

for filename in sys.argv[1:]:
    new_filename = filename + ".new"
    with open(filename, encoding="utf-8") as old_f, open(
        new_filename, "w", encoding="utf-8",
    ) as new_f:
        new_f.writelines(
            re.sub("^import T", "import numbers_parser.generated.T", line)
            for line in old_f
        )

    os.rename(new_filename, filename)
