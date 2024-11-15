import re
import sys
from subprocess import PIPE, Popen

# Code pattern in AArch64:
#
#         mov     w8, #325
#         strh    w8, [sp, #8]
#         str     x23, [sp]
#         adrp    x2, 2590 ; 0x14e0000
#         add     x2, x2, #3776 ; Objc cfstring ref: @"GETPIVOTDATA"
#
# TSCEFunction_GETPIVOTDATA::evaluateWithContext(...

# Additional required code pattern in AArch64 for Numbers 14.2:
#
#         mov     w0, #319
#         bl      TSCEFormulaCreationMagic::function_3arg(...
#         ;
#         ; Approx. 20 lines
#         ;
# TSCEFormulaCreationMagic::TEXTBETWEEN(...

if len(sys.argv) != 3:
    print(f"Usage: {sys.argv[0]} framework-file output.py", file=sys.stderr)
    sys.exit(1)

framework = sys.argv[1]
output_map = sys.argv[2]

if framework.endswith(".s"):
    with open(framework, "rb") as fh:
        disassembly = fh.readlines()
else:
    objdump = Popen(  # noqa: S603
        [  # noqa: S607
            "objdump",
            "--disassemble",
            "--no-addresses",
            "--no-print-imm-hex",
            "--no-show-raw-insn",
            "--macho",
            "--objc-meta-data",
            framework,
        ],
        stdout=PIPE,
    )
    cxxfilt = Popen(["c++filt"], stdin=objdump.stdout, stdout=PIPE)  # noqa: S603, S607
    objdump.stdout.close()
    disassembly = str(cxxfilt.communicate()[0]).split("\\n")

arg = None
line_count = 0
tsce_functions = {}
function_refs = {}

previous_line = ""
for line in disassembly:
    line = str(line).replace("\\t", " ")  # noqa: PLW2901
    if m := re.search(r"mov *w8, #(\d+)", line):
        arg = m.group(1)
        line_count = 0
        continue

    if arg is not None:
        line_count += 1
        if line_count > 30:
            arg = None
            line_count = 0

    if m := re.search(r'x2, x2.*Objc cfstring ref: @"([A-Z0-9\.]+)"', line):
        if arg is not None and line_count <= 8:
            func = m.group(1).replace("_", ".")
            print(f"Found cstring {func} = {arg}")
            function_refs[func] = arg
            arg = None
            line_count = 0
    elif m := re.search(r"bl *TSCEFormulaCreationMagic::function_3arg\(", line):
        if m := re.search(r"mov *w0, #(\d+)", previous_line):
            arg = m.group(1)
            line_count = 0
    elif (m := re.search(r"TSCEFormulaCreationMagic::(\w+)\(", line)) and arg is not None:
        func = m.group(1).replace("_", ".")
        print(f"Found TSCEFormulaCreationMagic {func} = {arg}")
        function_refs[func] = arg
        arg = None
        line_count = 0
    elif m := re.search(r"TSCEFunction_(\w+)::evaluateWithContext", line):
        func = m.group(1).replace("_", ".")
        print(f"Found TSCEFunction {func}")
        tsce_functions[func] = True

    previous_line = line

function_refs = dict(sorted(function_refs.items(), key=lambda x: int(x[1])))
with open(output_map, "w") as fh:
    fh.write("FUNCTION_MAP = {\n")
    if "ABS" not in function_refs:
        fh.write('    1: "ABS",\n')
    for func_name, func_id in function_refs.items():
        if func in tsce_functions:
            fh.write(f'    {func_id}: "{func_name}",\n')

    fh.write("}\n")
