import re
import sys

from subprocess import Popen, PIPE

# Code pattern in AArch64:
#
# 2824749   ac2c0c:   a8 28 80 52 mov w8, #325
# 2824750   ac2c10:   e8 13 00 79 strh    w8, [sp, #8]
# 2824751   ac2c14:   f7 03 00 f9 str x23, [sp]
# 2824752   ac2c18:   e2 50 00 d0 adrp    x2, 2590 ; 0x14e0000
# 2824753   ac2c1c:   42 00 3b 91 add x2, x2, #3776 ; Objc cfstring ref: @"GETPIVOTDATA"
#
# TSCEFunction_GETPIVOTDATA::evaluateWithContext(

if len(sys.argv) != 3:
    print(f"Usage: {sys.argv[0]} framework-file output.py", file=sys.stderr)
    sys.exit(1)

framework = sys.argv[1]
output_map = sys.argv[2]

objdump = Popen(
    [
        "objdump",
        "--disassemble",
        "--macho",
        "--objc-meta-data",
        framework,
    ],
    stdout=PIPE,
)
cxxfilt = Popen(["c++filt"], stdin=objdump.stdout, stdout=PIPE)
objdump.stdout.close()
disassembly = cxxfilt.communicate()[0]

arg = None
line_count = 0
tsce_functions = {}
cstring_refs = {}

for line in str(disassembly).split("\\n"):
    line = str(line).replace("\\t", " ")
    if m := re.search(r"mov *w8, #(\d+)", line):
        arg = m.group(1)
        continue

    if arg is not None:
        line_count += 1
        if line_count > 8:
            arg = None
            line_count = 0

    if m := re.search(r'x2, x2.*Objc cfstring ref: @"([A-Z0-9\.]+)"', line):
        if arg is not None and line_count <= 8:
            cstring_refs[m.group(1)] = arg
            arg = None
            line_count = 0

    if m := re.search(r"TSCEFunction_(\w+)::", line):
        tsce_functions[m.group(1)] = True

with open(output_map, "w") as fh:
    fh.write("FUNCTION_MAP = {\n")
    fh.write('    1: "ABS",\n')
    for func, id in cstring_refs.items():
        func_c = func.replace(".", "_")
        if func_c in tsce_functions:
            fh.write(f'    {id}: "{func}",\n')
        else:
            fh.write(f"    # {func_c} {func}\n")

    fh.write("}\n")
