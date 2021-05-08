import os
import re
import sys

for filename in sys.argv[1:]:
    new_filename = filename + ".new"
    old_f = open(filename, "r")
    new_f = open(new_filename, "w")
    for line in old_f.readlines():
      line = re.sub('^import T', 'import numbers_parser.generated.T', line)
      new_f.write(line)

    old_f.close()
    new_f.close()

    os.rename(new_filename, filename)
