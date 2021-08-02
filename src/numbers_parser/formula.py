import re
import warnings

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives

_FUNCTIONS = {
    1: "ABS",
    6: "ADDRESS",
    8: "AREAS",
    14: "AVEDEV",
    15: "AVERAGE",
    16: "AVERAGEA",
    17: "CEILING",
    18: "CHAR",
    20: "CLEAN",
    21: "CODE",
    22: "COLUMN",
    23: "COLUMNS",
    24: "COMBIN",
    25: "CONCATENATE",
    26: "CONFIDENCE",
    27: "CORREL",
    30: "COUNT",
    33: "COUNTIF",
    38: "COVAR",
    46: "DOLLAR",
    48: "EVEN",
    49: "EXACT",
    50: "EXP",
    51: "FACT",
    53: "FIND",
    54: "FIXED",
    55: "FLOOR",
    56: "FORECAST",
    58: "GCD",
    59: "HLOOKUP",
    62: "IF",
    63: "INDEX",
    65: "INT",
    66: "INTERCEPT",
    69: "ISBLANK",
    70: "ISERROR",
    71: "ISEVEN",
    75: "LCM",
    76: "LEFT",
    77: "LEN",
    78: "LN",
    79: "LOG",
    80: "LOG10",
    82: "LOWER",
    86: "MEDIAN",
    87: "MID",
    92: "MOD",
    95: "MROUND",
    97: "NOW",
    100: "ODD",
    104: "PI",
    106: "POISSONDIST",
    107: "POWER",
    113: "PRODUCT",
    114: "PROPER",
    116: "QUOTIENT",
    117: "RADIANS",
    118: "RAND",
    119: "RANDBETWEEN",
    122: "REPLACE",
    123: "REPT",
    124: "RIGHT",
    125: "ROMAN",
    126: "ROUND",
    127: "ROUNDDOWN",
    128: "ROUNDUP",
    129: "ROW",
    130: "ROWS",
    131: "SEARCH",
    133: "SIGN",
    139: "SQRT",
    145: "SUMIF",
    146: "SUMPRODUCT",
    147: "SUMSQ",
    149: "T",
    155: "TRIM",
    157: "TRUNC",
    165: "VLOOKUP",
    168: "SUM",
    198: "STANDARDIZE",
    213: "EXPONDIST",
    216: "SUMX2MY2",
    217: "SUMX2PY2",
    218: "SUMXMY2",
    219: "SQRTPI",
    220: "TRANSPOSE",
    221: "DEVSQ",
    222: "FREQUENCY",
    224: "FACTDOUBLE",
    227: "GAMMALN",
    229: "GAMMADIST",
    230: "GAMMAINV",
    234: "AVERAGEIF",
    240: "LOGNORMINV",
    242: "BINOMDIST",
    244: "FDIST",
    246: "CHIDIST",
    247: "CHITEST",
    250: "MULTINOMIAL",
    251: "CRITBINOM",
    256: "CHINV",
    257: "FINV",
    259: "BETAINV",
    262: "HARMEAN",
    263: "GEOMEAN",
    280: "INTERSECT.RANGES",
    286: "SERIESSUM",
    304: "ISNUMBER",
    305: "ISTEXT",
}


class Formula(list):
    def __init__(self, row_num, col_num):
        self._stack = []
        self.row = row_num
        self.col = col_num

    def __str__(self):
        return "".join(reversed(self._stack))

    def pop(self) -> str:
        return self._stack.pop()

    def popn(self, num_args: int) -> tuple:
        values = ()
        for i in range(num_args):
            values += (self._stack.pop(),)
        return values

    def push(self, val: str):
        self._stack.append(val)

    def function(self, num_args: int, node_index: int):
        if node_index not in _FUNCTIONS:
            _ = self.popn(num_args)
            warnings.warn(
                f"@[{self.row},{self.col}: function ID {node_index} is unsupported",
                UnsupportedWarning,
            )
            self.push(f"*UNSUPPORTED:{node_index}*")
            return
        else:
            func_name = _FUNCTIONS[node_index]
            if len(self._stack) < num_args:
                warnings.warn(
                    f"@[{self.row},{self.col}: stack to small for {func_name}",
                    UnsupportedWarning,
                )
                num_args = len(self._stack)
            args = self.popn(num_args)
            args = ",".join(reversed(args))
            self.push(f"{func_name}({args})")

    def array(self, num_rows: int, num_cols: int):
        # Excel array format:
        #     1-dimentional: {a,b,c,d}
        #     2-dimentional: {a,b;c,d}
        if num_rows == 1:
            args = self.popn(num_cols)
            args = ",".join(reversed(args))
            self.push(f"{{{args}}}")
        else:
            rows = []
            for row_num in range(num_rows):
                args = self.popn(num_cols)
                args = ",".join(reversed(args))
                rows.append(f"{args}")
            args = ";".join(reversed(rows))
            self.push(f"{{{args}}}")

    def list(self, num_args: int):
        args = self.popn(num_args)
        args = ",".join(reversed(args))
        self.push(f"({args})")

    def add(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def concat(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def div(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}/{arg2}")

    def equals(self):
        arg1, arg2 = self.popn(2)
        # TODO: arguments reversed?
        self.push(f"{arg2}={arg1}")

    def greater_than(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def greater_than_or_equal(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>={arg2}")

    def mul(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}*{arg2}")

    def negate(self):
        arg1 = self.pop()
        self.push(f"-{arg1}")

    def not_equals(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def power(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}^{arg2}")

    def range(self):
        arg2, arg1 = self.popn(2)
        func_range = "(" in arg1 or "(" in arg2
        if "::" in arg1 and not func_range:
            # Assumes references are not cross-table
            arg1_parts = arg1.split("::")
            arg2_parts = arg2.split("::")
            self.push(f"{arg1_parts[0]}::{arg1_parts[1]}:{arg2_parts[1]}")
        else:
            self.push(f"{arg1}:{arg2}")

    def sub(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")


class FormulaNode:
    def __init__(self, node_type, **kwargs):
        self.type = node_type
        for arg, val in kwargs.items():
            setattr(self, arg, val)


class TableFormulas:
    def __init__(self, model, table_id):
        self._model = model
        self._table_id = table_id
        self._formula_type_lookup = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

    def is_formula(self, row, col):
        if not (hasattr(self, "_formula_cells")):
            self._formula_cells = self._model.formula_cell_ranges(self._table_id)

        return (row, col) in self._formula_cells

    def is_error(self, row, col):
        if not (hasattr(self, "_error_cells")):
            self._error_cells = self._model.error_cell_ranges(self._table_id)
        return (row, col) in self._error_cells

    def formula(self, formula_key, row_num, col_num):  # noqa: C901
        if not (hasattr(self, "_ast")):
            self._ast = self._model.formula_ast(self._table_id)
        if formula_key not in self._ast:
            return "*INVALID KEY*"
        ast = self._ast[formula_key]
        formula = Formula(row_num, col_num)
        for node in ast:
            node_type = self._formula_type_lookup[node.AST_node_type]
            if node_type == "ADDITION_NODE":
                formula.add()
            elif node_type == "ARRAY_NODE":
                formula.array(node.AST_array_node_numRow, node.AST_array_node_numCol)
            elif node_type == "BEGIN_EMBEDDED_NODE_ARRAY":
                pass
            elif node_type == "BOOLEAN_NODE":
                formula.push(str(node.AST_boolean_node_boolean).upper())
            elif node_type == "CELL_REFERENCE_NODE":
                formula.push(self._model.node_to_cell_ref(row_num, col_num, node))
            elif node_type == "COLON_NODE":
                formula.range()
            elif node_type == "CONCATENATION_NODE":
                formula.concat()
            elif node_type == "DIVISION_NODE":
                formula.div()
            elif node_type == "END_THUNK_NODE":
                pass
            elif node_type == "EMPTY_ARGUMENT_NODE":
                formula.push("")
            elif node_type == "EQUAL_TO_NODE":
                formula.equals()
            elif node_type == "EQUAL_TO_NODE":
                formula.equals()
            elif node_type == "FUNCTION_NODE":
                formula.function(
                    node.AST_function_node_numArgs, node.AST_function_node_index
                )
            elif node_type == "GREATER_THAN_NODE":
                formula.greater_than()
            elif node_type == "GREATER_THAN_OR_EQUAL_TO_NODE":
                formula.greater_than_or_equal()
            elif node_type == "LIST_NODE":
                formula.list(node.AST_list_node_numArgs)
            elif node_type == "MULTIPLICATION_NODE":
                formula.mul()
            elif node_type == "NEGATION_NODE":
                formula.negate()
            elif node_type == "NOT_EQUAL_TO_NODE":
                formula.not_equals()
            elif node_type == "NUMBER_NODE":
                if node.AST_number_node_decimal_high == 0x3040000000000000:
                    # Integer: don't use decimals
                    formula.push(str(node.AST_number_node_decimal_low))
                else:
                    # TODO: detect when scientific notation is present
                    formula.push(number_to_str(node.AST_number_node_number))
            elif node_type == "POWER_NODE":
                formula.power()
            elif node_type == "PREPEND_WHITESPACE_NODE":
                # TODO: something with node.AST_whitespace
                pass
            elif node_type == "STRING_NODE":
                formula.push('"' + node.AST_string_node_string + '"')
            elif node_type == "SUBTRACTION_NODE":
                formula.sub()
            elif node_type == "TOKEN_NODE":
                if node.AST_token_node_boolean:
                    formula.push("TRUE")
                else:
                    formula.push("FALSE")
            else:
                warnings.warn(
                    f"@[{row_num},{col_num}: function node type {node_type} is unsupported",
                    UnsupportedWarning,
                )
                return str(node_type)

        return str(formula)


def number_to_str(v: int) -> str:
    v_str = repr(v)
    if "e" in v_str:
        number, exp = v_str.split("e")
        number = re.sub(r"[,-.]", "", number)
        zeroes = "0" * (abs(int(exp)) - 1)
        if v < 0:
            sign = "-"
        else:
            sign = ""
        if int(exp) > 0:
            return f"{sign}{number}{zeroes}"
        else:
            return f"{sign}0.{zeroes}{number}"
    else:
        return v_str
