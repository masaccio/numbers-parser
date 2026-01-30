import re
import sys
from subprocess import PIPE, Popen

from numbers_parser.generated.functionmap import FUNCTION_MAP as OLD_FUNC_ID_TO_NAME

OLD_FUNC_NAME_TO_ID = {v: k for k, v in OLD_FUNC_ID_TO_NAME.items()}
CFSTRING_MAX_LINES = 10
FORMULA_CREATION_MAX_LINES = 50

# Since Numbers 14.2, this code sequence indicates a formula node being created. The
# spreadsheet function is the class method name and the ID is loaded into register W0:
#
# TSCEFormulaCreationMagic::AND(TSCEFormulaCreator, TSCEFormulaCreator):
#        mov w0, #7
#         ;
#         ; Approx. 10-50 lines
#         ;
#         bl  TSCEFormulaCreationMagic::function_<n>arg(...
#
# Before this, the following code sequence could be used where CFString is used and
# with the name of the spreadsheet function just after loading an IS into register W8:
#
#         mov     w8, #325
#         strh    w8, [sp, #8]
#         str     x23, [sp]
#         adrp    x2, 2590 ; 0x14e0000
#         add     x2, x2, #3776 ; Objc cfstring ref: @"GETPIVOTDATA"
#


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
    cxxfilt = Popen(["c++filt"], stdin=objdump.stdout, stdout=PIPE)  # noqa: S607
    objdump.stdout.close()
    disassembly = str(cxxfilt.communicate()[0]).split("\\n")

cfstring_arg = None
cfstring_line_count = 0
cfstring_func_name_to_id = {}

formula_creation_arg = False
formula_creation_line_count = 0
formula_creation_name_to_id = {}

tsce_functions = {}

previous_line = ""
for line in disassembly:
    line = str(line).replace("\\t", " ")  # noqa: PLW2901

    # Potential start of code sequence 1 (Objc cfstring ref)
    if m := re.search(r"mov *w8, #(\d+)", line):
        cfstring_arg = m.group(1)
        cfstring_line_count = 0
        continue

    if cfstring_arg is not None:
        cfstring_line_count += 1
        if cfstring_line_count > CFSTRING_MAX_LINES:
            cfstring_arg = None
            cfstring_line_count = 0

    if formula_creation_arg is not None:
        formula_creation_line_count += 1
        if formula_creation_line_count > FORMULA_CREATION_MAX_LINES:
            formula_creation_arg = None
            formula_creation_line_count = 0

    # End of code sequence 1 (Objc cfstring ref)
    if cfstring_arg is not None and (
        m := re.search(r'x2, x2.*Objc cfstring ref: @"([A-Z0-9\.]+)"', line)
    ):
        func = m.group(1).replace("_", ".")
        print(f"Found cstring {func} = {cfstring_arg}")
        cfstring_func_name_to_id[func] = int(cfstring_arg)
        cfstring_arg = None
        cfstring_line_count = 0

    # Start of code sequence 2 (defining TSCEFormulaCreationMagic function)
    if m := re.search(r"^TSCEFormulaCreationMagic::(\w+)\(TSCEFormulaCreator", line):
        formula_creation_arg = m.group(1).replace("_", ".")
        formula_creation_line_count = 0

    # End of code sequence 2 (calling TSCEFormulaCreationMagic::function_<n>arg)
    if (
        formula_creation_arg is not None
        and re.search(r"bl *TSCEFormulaCreationMagic::function_[0-9]arg\(", line)
        and (m := re.search(r"mov *w0, #(\d+)", previous_line))
    ):
        print(f"Found TSCEFormulaCreationMagic {formula_creation_arg} = {m.group(1)}")
        formula_creation_name_to_id[formula_creation_arg] = int(m.group(1))
        formula_creation_arg = None
        formula_creation_line_count = 0

    if m := re.search(r"TSCEFunction_(\w+) evaluateForArgsWithContext", line):
        func = m.group(1).replace("_", ".")
        print(f"Found TSCEFunction {func}")
        tsce_functions[func] = True

    previous_line = line

# ID 1 is ABS() but the code sequence scanning finds multiple ID=1 functions
function_refs = {k: v for k, v in formula_creation_name_to_id.items() if v != 1}
function_refs.update(
    {k: v for k, v in cfstring_func_name_to_id.items() if k in tsce_functions and v != 1},
)
function_refs = dict(sorted(function_refs.items(), key=lambda x: int(x[1])))

for func_name, func_id in OLD_FUNC_NAME_TO_ID.items():
    if func_name not in function_refs:
        print(f"*** {func_name} has been removed (was ID {func_id})")
    elif function_refs[func_name] != func_id:
        print(
            f"*** {func_name} bad ID: is {function_refs[func_name]}, should be {func_id}",
        )

with open(output_map, "w") as fh:
    fh.write("FUNCTION_MAP = {\n")
    if "ABS" not in function_refs:
        fh.write('    1: "ABS",\n')
    for func_name, func_id in function_refs.items():
        fh.write(f'    {func_id}: "{func_name}",\n')

    fh.write("}\n")
