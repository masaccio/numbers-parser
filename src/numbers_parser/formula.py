import re
import warnings

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.functionmap import FUNCTION_MAP

from datetime import datetime, timedelta


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
        if node_index not in FUNCTION_MAP:
            _ = self.popn(num_args)
            warnings.warn(
                f"@[{self.row},{self.col}: function ID {node_index} is unsupported",
                UnsupportedWarning,
            )
            self.push(f"*UNSUPPORTED:{node_index}*")
            return
        else:
            func_name = FUNCTION_MAP[node_index]
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
            elif node_type == "APPEND_WHITESPACE_NODE":
                pass
            elif node_type == "ARRAY_NODE":
                formula.array(node.AST_array_node_numRow, node.AST_array_node_numCol)
            elif node_type == "BEGIN_EMBEDDED_NODE_ARRAY":
                pass
            elif node_type == "BOOLEAN_NODE":
                formula.push(str(node.AST_boolean_node_boolean).upper())
            elif node_type == "CELL_REFERENCE_NODE":
                formula.push(self._model.node_to_ref(row_num, col_num, node))
            elif node_type == "COLON_NODE":
                formula.range()
            elif node_type == "CONCATENATION_NODE":
                formula.concat()
            elif node_type == "DATE_NODE":
                dt = datetime(2001, 1, 1) + timedelta(
                    seconds=node.AST_date_node_dateNum
                )
                formula.push(f"DATE({dt.year},{dt.month},{dt.day})")
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
                pass

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
