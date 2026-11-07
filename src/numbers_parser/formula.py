"""Convert Numbers formula archives into readable formula strings."""

from __future__ import annotations

import re
import warnings
from datetime import datetime, timedelta
from typing import TYPE_CHECKING

from numbers_parser.constants import NUMBER_NODE_DECIMAL_HIGH
from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated.functionmap import FUNCTION_MAP

if TYPE_CHECKING:
    from numbers_parser.model import NumbersModel

FUNCTION_NAME_TO_ID = {v: k for k, v in FUNCTION_MAP.items()}


class Formula(list):
    """Convert a cell formula into readable text."""

    def __init__(
        self,
        model: NumbersModel,
        table_id: int,
        row: int,
        col: int,
    ) -> None:
        """Initialize the instance state."""
        self._stack = []
        self._model = model
        self._table_id = table_id
        self.row = row
        self.col = col

    def __str__(self) -> str:
        """Return the string representation."""
        return "".join(reversed([str(x) for x in self._stack]))

    def pop(self) -> str:
        """Pop the value."""
        return self._stack.pop()

    def popn(self, num_args: int) -> tuple[str, ...]:
        """Pop and return the requested number of values from the formula stack."""
        values = ()
        for _ in range(num_args):
            values += (self._stack.pop(),)
        return values

    def push(self, val: str) -> None:
        """Push the value."""
        self._stack.append(val)

    def add(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Add the value."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def array(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Build formula text for a one- or two-dimensional array."""
        node = args[2]
        num_rows = node.AST_array_node_numRow
        num_cols = node.AST_array_node_numCol
        if num_rows == 1:
            # 1-dimensional array: {a,b,c,d}
            args = self.popn(num_cols)
            args = ",".join(reversed(args))
            self.push(f"{{{args}}}")
        else:
            # 2-dimensional array: {a,b;c,d}
            rows = []
            for _row_num in range(num_rows):
                args = self.popn(num_cols)
                args = ",".join(reversed(args))
                rows.append(f"{args}")
            args = ";".join(reversed(rows))
            self.push(f"{{{args}}}")

    def boolean(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Push the formula text for a Boolean value onto the stack."""
        node = args[2]
        if node.HasField("AST_token_node_boolean"):
            self.push(str(node.AST_token_node_boolean).upper())
        else:
            self.push(str(node.AST_boolean_node_boolean).upper())

    def concat(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Join the top two formula values with the concatenation operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def date(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        # Date literals exported as DATE()
        """Format a Numbers date literal as a DATE() formula."""
        node = args[2]
        dt = datetime(2001, 1, 1) + timedelta(seconds=node.AST_date_node_dateNum)  # noqa: DTZ001
        self.push(f"DATE({dt.year},{dt.month},{dt.day})")

    def div(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Combine the top two formula values with the division operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}÷{arg2}")

    def empty(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Push an empty formula value onto the stack."""
        self.push("")

    def equals(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        # Arguments appear to be reversed
        """Compare the top two formula values for equality."""
        arg1, arg2 = self.popn(2)
        self.push(f"{arg2}={arg1}")

    def function(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Format a function node and its arguments as formula text."""
        node = args[2]
        num_args = node.AST_function_node_numArgs
        node_index = node.AST_function_node_index
        if node_index not in FUNCTION_MAP:
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{self.row},{self.col}]: function ID {node_index} is unsupported",
                UnsupportedWarning,
                stacklevel=2,
            )
            func_name = "UNDEFINED!"
        else:
            func_name = FUNCTION_MAP[node_index]

        if len(self._stack) < num_args:
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{self.row},{self.col}]: stack too small for {func_name}",
                UnsupportedWarning,
                stacklevel=2,
            )
            num_args = len(self._stack)

        args = self.popn(num_args)
        args = ",".join(reversed([str(x) for x in args]))
        self.push(f"{func_name}({args})")

    def greater_than(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Compare the top two formula values with the greater-than operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def greater_than_or_equal(
        self,
        *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the greater-than-or-equal operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≥{arg2}")

    def less_than(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Compare the top two formula values with the less-than operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}<{arg2}")

    def less_than_or_equal(
        self,
        *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the less-than-or-equal operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≤{arg2}")

    def list(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Format a list node and its values as formula text."""
        node = args[2]
        args = self.popn(node.AST_list_node_numArgs)
        args = ",".join(reversed([str(x) for x in args]))
        self.push(f"({args})")

    def mul(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Combine the top two formula values with the multiplication operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}×{arg2}")  # noqa: RUF001

    def negate(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Negate the top formula value."""
        arg1 = self.pop()
        self.push(f"-{arg1}")

    def not_equals(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Compare the top two formula values for inequality."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def number(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Format a numeric literal as formula text."""
        node = args[2]
        if node.AST_number_node_decimal_high == NUMBER_NODE_DECIMAL_HIGH:
            # Integer: don't use decimals
            self.push(str(node.AST_number_node_decimal_low))
        else:
            self.push(number_to_str(node.AST_number_node_number))

    def percent(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Append the percent operator to the top formula value."""
        arg1 = self.pop()
        self.push(f"{arg1}%")

    def power(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Combine the top two formula values with the exponentiation operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}^{arg2}")

    def range(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Format two cell references as a cell range."""
        arg2, arg1 = [str(x) for x in self.popn(2)]
        func_range = "(" in arg1 or "(" in arg2
        if "::" in arg1 and not func_range:
            # Assumes references are not cross-table
            arg1_parts = arg1.split("::")
            arg2_parts = arg2.split("::")
            self.push(f"{arg1_parts[0]}::{arg1_parts[1]}:{arg2_parts[1]}")
        else:
            self.push(f"{arg1}:{arg2}")

    def string(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Format a string literal, escaping embedded quotation marks."""
        node = args[2]
        # Numbers does not escape quotes in the AST; in the app, they are
        # doubled up just like in Excel
        value = node.AST_string_node_string.replace('"', '""')
        self.push(f'"{value}"')

    def sub(self, *_args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Combine the top two formula values with the subtraction operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")

    def xref(self, *args: int | TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive) -> None:
        """Push the cell reference for a formula node onto the stack."""
        (row, col, node) = args
        self.push(self._model.node_to_ref(self._table_id, row, col, node))


NODE_FUNCTION_MAP = {
    "ADDITION_NODE": "add",
    "APPEND_WHITESPACE_NODE": None,
    "ARRAY_NODE": "array",
    "BEGIN_EMBEDDED_NODE_ARRAY": None,
    "BEGIN_THUNK_NODE": None,
    "BOOLEAN_NODE": "boolean",
    "CELL_REFERENCE_NODE": "xref",
    "COLON_NODE": "range",
    "COLON_NODE_WITH_UIDS": "range",
    "COLON_TRACT_NODE": "xref",
    "CONCATENATION_NODE": "concat",
    "DATE_NODE": "date",
    "DIVISION_NODE": "div",
    "EMPTY_ARGUMENT_NODE": "empty",
    "END_THUNK_NODE": None,
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
    "PLUS_SIGN_NODE": None,
    "PREPEND_WHITESPACE_NODE": None,
    "STRING_NODE": "string",
    "SUBTRACTION_NODE": "sub",
    "TOKEN_NODE": "boolean",
}


class TableFormulas:
    """Convert table formula records into readable formulas."""

    def __init__(self, model: NumbersModel, table_id: int) -> None:
        """Initialize the instance state."""
        self._model = model
        self._table_id = table_id
        self._formula_type_lookup = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()  # noqa: SLF001
        }

    def formula(self, formula_key: int, row: int, col: int) -> str:
        """Return the formula text for a cell in the table."""
        all_formulas = self._model.formula_ast(self._table_id)
        if formula_key not in all_formulas:
            table_name = self._model.table_name(self._table_id)
            warnings.warn(
                f"{table_name}@[{row},{col}]: key #{formula_key} not found",
                UnsupportedWarning,
                stacklevel=2,
            )
            return "INVALID_KEY!(" + str(formula_key) + ")"

        formula = Formula(self._model, self._table_id, row, col)
        for node in all_formulas[formula_key]:
            node_type = self._formula_type_lookup[node.AST_node_type]
            if node_type == "REFERENCE_ERROR_WITH_UIDS":
                formula.push("#REF!")
            elif node_type not in NODE_FUNCTION_MAP:
                table_name = self._model.table_name(self._table_id)
                warnings.warn(
                    f"{table_name}@[{row},{col}]: node type {node_type} is unsupported",
                    UnsupportedWarning,
                    stacklevel=2,
                )
            elif NODE_FUNCTION_MAP[node_type] is not None:
                func = getattr(formula, NODE_FUNCTION_MAP[node_type])
                func(row, col, node)

        return str(formula)


def number_to_str(v: int) -> str:
    """Format a float as a string."""
    # Number is never negative; formula will use NEGATION_NODE
    v_str = repr(v)
    if "e" in v_str:
        number, exp = v_str.split("e")
        number = re.sub(r"[,-]", "", number)
        num_zeroes = abs(int(exp)) - 1
        if "." in number:
            num_zeroes -= len(number.split(".")[1]) - 1
            number = number.replace(".", "")
        zeroes = "0" * num_zeroes
        if int(exp) > 0:
            return f"{number}{zeroes}"
        return f"0.{zeroes}{number}"
    return v_str
