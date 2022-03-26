import re
import warnings

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.functionmap import FUNCTION_MAP

from datetime import datetime, timedelta


class Formula(list):
    def __init__(self, model, table_id, row_num, col_num):
        self._stack = []
        self._model = model
        self._table_id = table_id
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

    def add(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def array(self, *args):
        node = args[2]
        num_rows = node.AST_array_node_numRow
        num_cols = node.AST_array_node_numCol
        if num_rows == 1:
            # 1-dimentional array: {a,b,c,d}
            args = self.popn(num_cols)
            args = ",".join(reversed(args))
            self.push(f"{{{args}}}")
        else:
            # 2-dimentional array: {a,b;c,d}
            rows = []
            for row_num in range(num_rows):
                args = self.popn(num_cols)
                args = ",".join(reversed(args))
                rows.append(f"{args}")
            args = ";".join(reversed(rows))
            self.push(f"{{{args}}}")

    def boolean(self, *args):
        node = args[2]
        if node.HasField("AST_token_node_boolean"):
            self.push(str(node.AST_token_node_boolean).upper())
        else:
            self.push(str(node.AST_boolean_node_boolean).upper())

    def concat(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def date(self, *args):
        # Date literals exported as DATE()
        node = args[2]
        dt = datetime(2001, 1, 1) + timedelta(seconds=node.AST_date_node_dateNum)
        self.push(f"DATE({dt.year},{dt.month},{dt.day})")

    def div(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}÷{arg2}")

    def empty(self, *args):
        self.push("")

    def equals(self, *args):
        # Arguments appear to be reversed
        arg1, arg2 = self.popn(2)
        self.push(f"{arg2}={arg1}")

    def function(self, *args):
        node = args[2]
        num_args = node.AST_function_node_numArgs
        node_index = node.AST_function_node_index
        if node_index not in FUNCTION_MAP:  # pragma: no cover
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{self.row},{self.col}]: function ID {node_index} is unsupported",
                UnsupportedWarning,
            )
            func_name = "UNDEFINED!"
        else:
            func_name = FUNCTION_MAP[node_index]

        if len(self._stack) < num_args:  # pragma: no cover
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{self.row},{self.col}]: stack to small for {func_name}",
                UnsupportedWarning,
            )
            num_args = len(self._stack)

        args = self.popn(num_args)
        args = ",".join(reversed(args))
        self.push(f"{func_name}({args})")

    def greater_than(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def greater_than_or_equal(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≥{arg2}")

    def less_than(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}<{arg2}")

    def less_than_or_equal(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≤{arg2}")

    def list(self, *args):
        node = args[2]
        args = self.popn(node.AST_list_node_numArgs)
        args = ",".join(reversed(args))
        self.push(f"({args})")

    def mul(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}×{arg2}")

    def negate(self, *args):
        arg1 = self.pop()
        self.push(f"-{arg1}")

    def not_equals(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def number(self, *args):
        node = args[2]
        if node.AST_number_node_decimal_high == 0x3040000000000000:
            # Integer: don't use decimals
            self.push(str(node.AST_number_node_decimal_low))
        else:
            # TODO: detect when scientific notation is present
            self.push(number_to_str(node.AST_number_node_number))

    def percent(self, *args):
        arg1 = self.pop()
        self.push(f"{arg1}%")

    def power(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}^{arg2}")

    def range(self, *args):
        arg2, arg1 = self.popn(2)
        func_range = "(" in arg1 or "(" in arg2
        if "::" in arg1 and not func_range:
            # Assumes references are not cross-table
            arg1_parts = arg1.split("::")
            arg2_parts = arg2.split("::")
            self.push(f"{arg1_parts[0]}::{arg1_parts[1]}:{arg2_parts[1]}")
        else:
            self.push(f"{arg1}:{arg2}")

    def string(self, *args):
        node = args[2]
        self.push('"' + node.AST_string_node_string + '"')

    def sub(self, *args):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")

    def unsupported(self, *args):
        pass

    def xref(self, *args):
        (row_num, col_num, node) = args
        self.push(self._model.node_to_ref(row_num, col_num, node))


NODE_FUNCTION_MAP = {
    "ADDITION_NODE": "add",
    "APPEND_WHITESPACE_NODE": None,
    "ARRAY_NODE": "array",
    "BEGIN_EMBEDDED_NODE_ARRAY": None,
    "BOOLEAN_NODE": "boolean",
    "CELL_REFERENCE_NODE": "xref",
    "COLON_NODE": "range",
    # "COLON_NODE_WITH_UIDS": "unsupported",
    "CONCATENATION_NODE": "concat",
    "DATE_NODE": "date",
    "DIVISION_NODE": "div",
    "END_THUNK_NODE": None,
    "EMPTY_ARGUMENT_NODE": "empty",
    "EQUAL_TO_NODE": "equals",
    "FUNCTION_NODE": "function",
    "GREATER_THAN_NODE": "greater_than",
    "GREATER_THAN_OR_EQUAL_TO_NODE": "greater_than_or_equal",
    "LESS_THAN_NODE": "less_than",
    "LESS_THAN_OR_EQUAL_TO_NODE": "less_than_or_equal",
    "LIST_NODE": "list",
    "MULTIPLICATION_NODE": "mul",
    "NEGATION_NODE": "negate",
    "NOT_EQUAL_TO_NODE": "not_equals",
    "NUMBER_NODE": "number",
    "PERCENT_NODE": "percent",
    "POWER_NODE": "power",
    "PREPEND_WHITESPACE_NODE": None,
    "STRING_NODE": "string",
    "SUBTRACTION_NODE": "sub",
    "TOKEN_NODE": "boolean",
}


class TableFormulas:
    def __init__(self, model, table_id):
        self._model = model
        self._table_id = table_id
        self._formula_type_lookup = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

    def is_formula(self, row, col):
        return (row, col) in self._model.formula_cell_ranges(self._table_id)

    def formula(self, formula_key, row_num, col_num):
        all_formulas = self._model.formula_ast(self._table_id)
        if formula_key not in all_formulas:  # pragma: no cover
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{row_num},{col_num}]: key #{formula_key} not found",
                UnsupportedWarning,
            )
            return "INVALID_KEY!(" + str(formula_key) + ")"

        formula = Formula(self._model, self._table_id, row_num, col_num)
        for node in all_formulas[formula_key]:
            node_type = self._formula_type_lookup[node.AST_node_type]
            if node_type not in NODE_FUNCTION_MAP:  # pragma: no cover
                table_name = self._model.table_name(self._table_id)
                warnings.warn(
                    f"{table_name}@[{row_num},{col_num}]: node type {node_type} is unsupported",
                    UnsupportedWarning,
                )
                pass
            elif NODE_FUNCTION_MAP[node_type] is not None:
                func = getattr(formula, NODE_FUNCTION_MAP[node_type])
                func(row_num, col_num, node)

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
