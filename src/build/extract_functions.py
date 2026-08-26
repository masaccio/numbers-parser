import re
import sys
from subprocess import PIPE, Popen

from numbers_parser.generated.functionmap import FUNCTION_MAP

# Since Numbers 14.x, code sequences have become less and less reliable. Since Numbers function
# nodes will always be backwards-compatible, we can use the old IDs and only flag up those
# we are sure are new and missing. This metadata sequence which is definitive:
#
#   00000000015798f0 0x1656b70 _OBJC_CLASS_$_TSCEFunction_DUR2MILLISECONDS
#      isa        0x1656b48 _OBJC_METACLASS_$_TSCEFunction_DUR2MILLISECONDS
#      superclass 0x166f3a0 _OBJC_CLASS_$_TSCEFunctionNode
#
# This code sequence is also definitive but does not represent all functions:
#
#     TSCEFormulaCreationMagic::STRIPDURATION(TSCEFormulaCreator):
#         sub    sp, sp, #48
#         stp    x20, x19, [sp, #16]
#         stp    x29, x30, [sp, #32]
#         add    x29, sp, #32
#         mov    x20, x8
#         ldr    x0, [x0]
#         bl    0xf6d65c ; symbol stub for: _objc_retainBlock
#         mov    x19, x0
#         str    x0, [sp, #8]
#         add    x1, sp, #8
#         mov    x8, x20
#         mov    w0, #278
#         bl    TSCEFormulaCreationMagic::function_1arg(TSCEFunctionIndex, TSCEFormulaCreator)

TSCE_FORMULA_IGNORE_PREFIXES = (".", "op.", "RANGE.TRACKING", "SHOW")

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

formula_creation_name = None
formula_creation_name_to_id = {}
tsce_functions = {}

previous_line = ""
for line in disassembly:
    if isinstance(line, bytes):
        line = line.decode(encoding="latin1")  # noqa: PLW2901
    line = line.replace("\\t", " ").strip()  # noqa: PLW2901

    if m := re.search(r"^TSCEFormulaCreationMagic::(.*?)\(", line):
        formula_creation_name = m.group(1).replace("_", ".")
        if formula_creation_name.startswith(TSCE_FORMULA_IGNORE_PREFIXES):
            formula_creation_name = None

    if formula_creation_name and (m := re.search(r"bl\s+TSCEFormulaCreationMagic::", line)):
        if m := re.search(r"mov\s+w\d, #(\d+)", previous_line):
            arg = m.group(1)
            print(f"Found TSCEFormulaCreationMagic for '{formula_creation_name}' with ID #{arg}")
            formula_creation_name_to_id[formula_creation_name] = int(arg)
        formula_creation_name = None

    if m := re.search(r"\bisa\b.*_OBJC_METACLASS_\$_TSCEFunction_(.*)", line):
        formula_creation_name = m.group(1).replace("_", ".")
        if not formula_creation_name.startswith(("..", "op.")):
            print(f"Found TSCEFunction '{formula_creation_name}'")
            tsce_functions[formula_creation_name] = True

    previous_line = line

function_refs = {v: k for k, v in formula_creation_name_to_id.items()}

for func_id, old_func_name in FUNCTION_MAP.items():
    if func_id not in function_refs:
        print(f"Retained {old_func_name} with ID #{func_id}")
        function_refs[func_id] = old_func_name
    else:
        func_name = function_refs[func_id]
        if function_refs[func_id] != old_func_name:
            print(f"*** {func_id} mismatch: was '{old_func_name}', now {func_name}")
            function_refs[func_id] = old_func_name

function_refs = dict(sorted(function_refs.items()))

for func_id, func_name in function_refs.items():
    if func_id not in FUNCTION_MAP:
        print(f"*** {func_id} is new: function is '{func_name}'")

with open(output_map, "w") as fh:
    fh.write("FUNCTION_MAP = {\n")
    fh.writelines(
        f'    {func_id}: "{func_name}",\n' for func_id, func_name in function_refs.items()
    )

    fh.write("}\n")
