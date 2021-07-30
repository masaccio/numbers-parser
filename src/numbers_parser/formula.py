import re
import warnings

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives

_FUNCTIONS = {
    1: {"func": "ABS", "nargs": 1},
    15: {"func": "AVERAGE", "nargs": 1},
    17: {"func": "CEILING", "nargs": [1, 2]},
    18: {"func": "CHAR", "nargs": 1},
    20: {"func": "CLEAN", "nargs": 1},
    21: {"func": "CODE", "nargs": 1},
    22: {"func": "COLUMN", "nargs": 0},
    24: {"func": "COMBIN", "nargs": 2},
    25: {"func": "CONCATENATE", "nargs": "*"},
    46: {"func": "DOLLAR", "nargs": 1},
    48: {"func": "EVEN", "nargs": 1},
    49: {"func": "EXACT", "nargs": 2},
    50: {"func": "EXP", "nargs": 1},
    51: {"func": "FACT", "nargs": 1},
    53: {"func": "FIND", "nargs": 2},
    54: {"func": "FIXED", "nargs": [1, 2, 3]},
    55: {"func": "FLOOR", "nargs": 2},
    58: {"func": "GCD", "nargs": "*"},
    62: {"func": "IF", "nargs": 3},
    65: {"func": "INT", "nargs": 1},
    69: {"func": "ISBLANK", "nargs": 1},
    70: {"func": "ISERROR", "nargs": 1},
    71: {"func": "ISEVEN", "nargs": 1},
    75: {"func": "LCM", "nargs": "*"},
    76: {"func": "LEFT", "nargs": [1, 2]},
    77: {"func": "LEN", "nargs": [1, 2]},
    78: {"func": "LN", "nargs": 1},
    79: {"func": "LOG", "nargs": [1, 2]},
    80: {"func": "LOG10", "nargs": 1},
    82: {"func": "LOWER", "nargs": 1},
    86: {"func": "MEDIAN", "nargs": 1},
    87: {"func": "MID", "nargs": [2, 3]},
    92: {"func": "MOD", "nargs": 2},
    95: {"func": "MROUND", "nargs": 2},
    97: {"func": "NOW", "nargs": 0},
    100: {"func": "ODD", "nargs": 1},
    104: {"func": "PI", "nargs": 0},
    107: {"func": "POWER", "nargs": 2},
    113: {"func": "PRODUCT", "nargs": "*"},
    114: {"func": "PROPER", "nargs": 0},
    116: {"func": "QUOTIENT", "nargs": 2},
    117: {"func": "RADIANS", "nargs": 1},
    118: {"func": "RAND", "nargs": 0},
    119: {"func": "RANDBETWEEN", "nargs": 2},
    122: {"func": "REPLACE", "nargs": 4},
    123: {"func": "REPT", "nargs": 2},
    124: {"func": "RIGHT", "nargs": [1, 2]},
    125: {"func": "ROMAN", "nargs": 2},
    126: {"func": "ROUND", "nargs": 2},
    127: {"func": "ROUNDDOWN", "nargs": 2},
    128: {"func": "ROUNDUP", "nargs": 2},
    131: {"func": "SEARCH", "nargs": [2, 3]},
    133: {"func": "SIGN", "nargs": 1},
    139: {"func": "SQRT", "nargs": 1},
    145: {"func": "SUMIF", "nargs": [2, 3]},
    146: {"func": "SUMPRODUCT", "nargs": 2},
    147: {"func": "SUMSQ", "nargs": "*"},
    149: {"func": "T", "nargs": 1},
    155: {"func": "TRIM", "nargs": 1},
    157: {"func": "TRUNC", "nargs": 1},
    168: {"func": "SUM", "nargs": "*"},
    216: {"func": "SUMX2MY2", "nargs": 2},
    217: {"func": "SUMX2PY2", "nargs": 2},
    218: {"func": "SUMXMY2", "nargs": 2},
    219: {"func": "SQRTPI", "nargs": 1},
    224: {"func": "FACTDOUBLE", "nargs": 1},
    250: {"func": "MULTINOMIAL", "nargs": [2, 3, 4]},
    286: {"func": "SERIESSUM", "nargs": 4},
    304: {"func": "ISNUMBER", "nargs": 1},
    305: {"func": "ISTEXT", "nargs": 1},
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

        f_nargs = _FUNCTIONS[node_index]["nargs"]
        f_func = _FUNCTIONS[node_index]["func"]
        if isinstance(f_nargs, str) and f_nargs == "*":
            args = self.popn(num_args)
            args = ",".join(reversed(args))
            self.push(f"{f_func}({args})")
        elif isinstance(f_nargs, list) and num_args in f_nargs:
            args = self.popn(num_args)
            args = ",".join(reversed(args))
            self.push(f"{f_func}({args})")
        elif f_nargs == 0:
            self.push(f"{f_func}()")
        elif f_nargs != num_args:
            warnings.warn(
                f"@[{self.row},{self.col}: arg count mismatch {f_func}: {num_args} vs {f_nargs}",
                UnsupportedWarning,
            )
            self.push(f"*ARG-ERROR:{f_func}:{num_args}:{f_nargs}*")
        else:
            args = self.popn(num_args)
            args = ",".join(reversed(args))
            self.push(f"{f_func}({args})")

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
        self.push(f"{arg1}={arg2}")

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
        if "::" in arg1:
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
                formula.push("(")
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
            else:
                warnings.warn(
                    f"@[{self.row},{self.col}: function node type {node_type} is unsupported",
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
