import logging
import warnings

from numbers_parser.exceptions import UnsupportedWarning

# from numbers_parser.document import Table

from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives

_FUNCTIONS = {
    15: {"func": "AVERAGE", "nargs": 1},
    18: {"func": "CHAR", "nargs": 1},
    22: {"func": "COLUMN", "nargs": 0},
    20: {"func": "CLEAN", "nargs": 1},
    25: {"func": "CODE", "nargs": 1},
    25: {"func": "CONCATENATE", "nargs": "*"},
    46: {"func": "DOLLAR", "nargs": 1},
    49: {"func": "EXACT", "nargs": 2},
    # 22: {"func": "FIXED", "nargs": 1},
    # 22: {"func": "LOWER", "nargs": 1},
    # 22: {"func": "PROPER", "nargs": 1},
    # 22: {"func": "REPLACE", "nargs": 1},
    53: {"func": "FIND", "nargs": 2},
    54: {"func": "FIXED", "nargs": [1, 2, 3]},
    62: {"func": "IF", "nargs": 3},
    69: {"func": "ISBLANK", "nargs": 1},
    70: {"func": "ISERROR", "nargs": 1},
    71: {"func": "ISEVEN", "nargs": 1},
    76: {"func": "LEFT", "nargs": [1, 2]},
    77: {"func": "LEN", "nargs": [1, 2]},
    82: {"func": "LOWER", "nargs": 1},
    86: {"func": "MEDIAN", "nargs": 1},
    87: {"func": "MID", "nargs": [2, 3]},
    97: {"func": "NOW", "nargs": 0},
    114: {"func": "PROPER", "nargs": 0},
    122: {"func": "REPLACE", "nargs": 2},
    123: {"func": "REPT", "nargs": 2},
    124: {"func": "RIGHT", "nargs": 2},
    149: {"func": "T", "nargs": 1},
    155: {"func": "TRIM", "nargs": 1},
    168: {"func": "SUM", "nargs": 1},
    304: {"func": "ISNUMBER", "nargs": 1},
    305: {"func": "ISTEXT", "nargs": 1},
}


class Formula(list):
    def __init__(self, row_num, col_num):
        self._stack = []
        self.row = row_num
        self.col = col_num

    def pop(self) -> str:
        return self._stack.pop()

    def popn(self, num_args: int) -> tuple:
        values = ()
        for i in range(num_args):
            values += (self._stack.pop(),)
        return values

    def push(self, val: str):
        self._stack.append(val)

    def __str__(self):
        return "//".join(self._stack)

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

    def equals(self):
        arg1, arg2 = self.popn(2)
        # TODO: arguments reversed?
        self.push(f"{arg1}={arg2}")

    def not_equals(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def greater_than(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def range(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}:{arg2}")

    def concat(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def add(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def sub(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")

    def div(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}/{arg2}")

    def mul(self):
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}*{arg2}")


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
        ast = self._ast[formula_key]
        formula = Formula(row_num, col_num)
        for node in ast:
            node_type = self._formula_type_lookup[node.AST_node_type]
            if node_type == "CELL_REFERENCE_NODE":
                formula.push(self._model.node_to_cell_ref(row_num, col_num, node))
            elif node_type == "NUMBER_NODE":
                # node.AST_number_node_decimal_high
                # node.AST_number_node_decimal_low
                if node.AST_number_node_decimal_high == 0x3040000000000000:
                    formula.push(str(node.AST_number_node_decimal_low))
                else:
                    formula.push(str(node.AST_number_node_number))
            elif node_type == "FUNCTION_NODE":
                formula.function(
                    node.AST_function_node_numArgs, node.AST_function_node_index
                )
            elif node_type == "STRING_NODE":
                formula.push('"' + node.AST_string_node_string + '"')
            elif node_type == "PREPEND_WHITESPACE_NODE":
                # TODO: something with node.AST_whitespace
                pass
            elif node_type == "EQUAL_TO_NODE":
                formula.equals()
            elif node_type == "ADDITION_NODE":
                formula.add()
            elif node_type == "BEGIN_EMBEDDED_NODE_ARRAY":
                formula.push("(")
            elif node_type == "BOOLEAN_NODE":
                pass
            elif node_type == "COLON_NODE":
                formula.range()
            elif node_type == "CONCATENATION_NODE":
                formula.concat()
            elif node_type == "DIVISION_NODE":
                formula.div()
            elif node_type == "END_THUNK_NODE":
                pass
            elif node_type == "EQUAL_TO_NODE":
                formula.equals()
            elif node_type == "GREATER_THAN_NODE":
                formula.greater_than()
            elif node_type == "MULTIPLICATION_NODE":
                formula.mul()
            elif node_type == "NOT_EQUAL_TO_NODE":
                formula.not_equals()
            elif node_type == "SUBTRACTION_NODE":
                formula.sub()
            else:
                return "FUNCTION_UNSUPPORTED:" + str(node_type)

        logging.debug(
            "%s@[%d,%d]: formula_key=%d, formula=%s",
            self._table_id,
            row_num,
            col_num,
            formula_key,
            formula,
        )

        return str(formula)
