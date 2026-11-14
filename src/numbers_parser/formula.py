"""Convert Numbers formula archives into readable formula strings."""

from __future__ import annotations

import re
import warnings
from datetime import datetime, timedelta
from typing import TYPE_CHECKING, cast

from numbers_parser.constants import NUMBER_NODE_DECIMAL_HIGH
from numbers_parser.exceptions import UnsupportedError, UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.xrefs import CellRange

if TYPE_CHECKING:
    from numbers_parser.model import NumbersModel

FUNCTION_NAME_TO_ID = {v: k for k, v in FUNCTION_MAP.items()}


class Formula:
    """Convert a cell formula into readable text."""

    def __init__(
        self,
        model: NumbersModel,
        table_id: int,
        row: int,
        col: int,
    ) -> None:
        """Initialize the instance state."""
        self._stack: list[str | CellRange] = []
        self._model = model
        self._table_id = table_id
        self.row = row
        self.col = col

    def __str__(self) -> str:
        """Return the string representation."""
        return "".join(reversed([str(x) for x in self._stack]))

    def pop(self) -> str | CellRange:
        """Pop the value."""
        return self._stack.pop()

    def popn(self, num_args: int) -> tuple[str | CellRange, ...]:
        """Pop and return the requested number of values from the formula stack."""
        values: tuple[str | CellRange, ...] = ()
        for _ in range(num_args):
            values += (self._stack.pop(),)
        return values

    def push(self, val: str | CellRange) -> None:
        """Push the value."""
        self._stack.append(val)

    def add(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Add the value."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def array(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Build formula text for a one- or two-dimensional array."""
        num_rows = node.AST_array_node_numRow
        num_cols = node.AST_array_node_numCol
        if num_rows == 1:
            # 1-dimensional array: {a,b,c,d}
            pop_args = self.popn(num_cols)
            joined_args = ",".join(str(value) for value in reversed(pop_args))
            self.push(f"{{{joined_args}}}")
        else:
            # 2-dimensional array: {a,b;c,d}
            rows = []
            for _row_num in range(num_rows):
                pop_args = self.popn(num_cols)
                joined_args = ",".join(str(value) for value in reversed(pop_args))
                rows.append(f"{joined_args}")
            joined_rows = ";".join(reversed(rows))
            self.push(f"{{{joined_rows}}}")

    def boolean(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Push the formula text for a Boolean value onto the stack."""
        if node.HasField("AST_token_node_boolean"):
            self.push(str(node.AST_token_node_boolean).upper())
        else:
            self.push(str(node.AST_boolean_node_boolean).upper())

    def concat(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Join the top two formula values with the concatenation operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def date(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        # Date literals exported as DATE()
        """Format a Numbers date literal as a DATE() formula."""
        dt = datetime(2001, 1, 1) + timedelta(seconds=node.AST_date_node_dateNum)  # noqa: DTZ001
        self.push(f"DATE({dt.year},{dt.month},{dt.day})")

    def div(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Combine the top two formula values with the division operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}÷{arg2}")

    def empty(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Push an empty formula value onto the stack."""
        self.push("")

    def equals(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        # Arguments appear to be reversed
        """Compare the top two formula values for equality."""
        arg1, arg2 = self.popn(2)
        self.push(f"{arg2}={arg1}")

    def function(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Format a function node and its arguments as formula text."""
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

        pop_args = self.popn(num_args)
        joined_args = ",".join(reversed([str(x) for x in pop_args]))
        self.push(f"{func_name}({joined_args})")

    def greater_than(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the greater-than operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def greater_than_or_equal(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the greater-than-or-equal operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≥{arg2}")

    def less_than(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the less-than operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}<{arg2}")

    def less_than_or_equal(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values with the less-than-or-equal operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≤{arg2}")

    def list(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Format a list node and its values as formula text."""
        pop_args = self.popn(node.AST_list_node_numArgs)
        joined_args = ",".join(reversed([str(x) for x in pop_args]))
        self.push(f"({joined_args})")

    def mul(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Combine the top two formula values with the multiplication operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}×{arg2}")  # noqa: RUF001

    def negate(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Negate the top formula value."""
        arg1 = self.pop()
        self.push(f"-{arg1}")

    def not_equals(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Compare the top two formula values for inequality."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def number(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Format a numeric literal as formula text."""
        if node.AST_number_node_decimal_high == NUMBER_NODE_DECIMAL_HIGH:
            # Integer: don't use decimals
            self.push(str(node.AST_number_node_decimal_low))
        else:
            self.push(number_to_str(node.AST_number_node_number))

    def percent(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Append the percent operator to the top formula value."""
        arg1 = self.pop()
        self.push(f"{arg1}%")

    def power(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Combine the top two formula values with the exponentiation operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}^{arg2}")

    def range(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Format two cell references as a cell range."""
        end, start = self.popn(2)
        if isinstance(start, CellRange) and isinstance(end, CellRange):
            if start._ref_error or end._ref_error or start.to_table_id != end.to_table_id:
                msg = "range endpoints target different or unavailable tables"
                raise UnsupportedError(msg)
            left, right = str(start), str(end)
            left_prefix, left_separator, left_cell = left.rpartition("::")
            right_prefix, right_separator, right_cell = right.rpartition("::")
            if left_separator or right_separator:
                if not left_separator or not right_separator or left_prefix != right_prefix:
                    msg = "range endpoint names cannot be combined safely"
                    raise UnsupportedError(msg)
                self.push(f"{left_prefix}::{left_cell}:{right_cell}")
                return
        self.push(f"{start}:{end}")

    def string(
        self,
        _row: int,
        _col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Format a string literal, escaping embedded quotation marks."""
        # Numbers does not escape quotes in the AST; in the app, they are
        # doubled up just like in Excel
        value = node.AST_string_node_string.replace('"', '""')
        self.push(f'"{value}"')

    def sub(
        self,
        _row: int,
        _col: int,
        _node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Combine the top two formula values with the subtraction operator."""
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")

    def xref(
        self,
        row: int,
        col: int,
        node: TSCEArchives.ASTNodeArrayArchive.ASTNodeArchive,
    ) -> None:
        """Push the cell reference for a formula node onto the stack."""
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
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()  # type: ignore[attr-defined] # noqa: SLF001
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
                func_name = cast("str", NODE_FUNCTION_MAP[node_type])
                func = getattr(formula, func_name)
                func(row, col, node)

        return str(formula)


def number_to_str(v: float) -> str:
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
