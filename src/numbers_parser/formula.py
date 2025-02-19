import math
import re
import warnings
from datetime import datetime, timedelta

from numbers_parser.constants import DECIMAL128_BIAS, OPERATOR_PRECEDENCE
from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.generated.TSCEArchives_pb2 import ASTNodeArrayArchive
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.tokenizer import Token, Tokenizer, parse_numbers_range
from numbers_parser.xrefs import CellRange, CellRangeType

FUNCTION_NAME_TO_ID = {v: k for k, v in FUNCTION_MAP.items()}

OPERATOR_MAP = str.maketrans({"×": "*", "÷": "/", "≥": ">=", "≤": "<=", "≠": "<>"})


OPERATOR_INFIX_MAP = {
    "=": "EQUAL_TO_NODE",
    "+": "ADDITION_NODE",
    "-": "SUBTRACTION_NODE",
    "*": "MULTIPLICATION_NODE",
    "/": "DIVISION_NODE",
    "&": "CONCATENATION_NODE",
    "^": "POWER_NODE",
    "==": "EQUAL_TO_NODE",
    "<>": "NOT_EQUAL_TO_NODE",
    "<": "LESS_THAN_NODE",
    ">": "GREATER_THAN_NODE",
    "<=": "LESS_THAN_OR_EQUAL_TO_NODE",
    ">=": "GREATER_THAN_OR_EQUAL_TO_NODE",
}

OPERAND_ARCHIVE_MAP = {
    Token.RANGE: "range_archive",
    Token.NUMBER: "number_archive",
    Token.TEXT: "text_archive",
    Token.LOGICAL: "logical_archive",
    Token.ERROR: "error",
}

# TODO: Understand what the frozen stick bits do!
FROZEN_STICKY_BIT_MAP = {
    (False, False, False, False): None,
    (False, True, False, False): (True, False, False, False),
    (False, False, False, True): (False, False, True, False),
    (False, True, False, True): (True, False, True, False),
    (False, False, True, False): None,
    (False, False, True, True): (False, False, True, False),
    (False, True, True, False): (True, False, False, False),
    (False, True, True, True): (True, False, True, False),
    (True, False, False, False): None,
    (True, True, False, False): (True, False, False, False),
    (True, False, False, True): (False, False, True, False),
    (True, True, False, True): (True, False, True, False),
    (True, False, True, False): None,
    (True, True, True, False): (True, False, False, False),
    (True, False, True, True): (False, False, True, False),
    (True, True, True, True): (True, False, True, False),
}


class Formula(list):
    def __init__(self, model, table_id, row, col) -> None:
        self._stack = []
        self._model = model
        self._table_id = table_id
        self.row = row
        self.col = col

    @classmethod
    def from_str(cls, model, table_id, row, col, formula_str) -> int:
        """
        Create a new formula by parsing a formula string and
        return the allocated formula ID.
        """
        formula = cls(model, table_id, row, col)
        formula._tokens = cls.formula_tokens(formula_str)

        model._formulas.add_table(table_id)
        formula_attrs = {"AST_node_array": {"AST_node": []}}
        ast_node = formula_attrs["AST_node_array"]["AST_node"]

        for token in formula._tokens:
            if token.type == Token.FUNC and token.subtype == Token.OPEN:
                if token.value not in FUNCTION_NAME_TO_ID:
                    table_name = model.table_name(table_id)
                    cell_ref = f"{table_name}@[{row},{col}]"
                    warnings.warn(
                        f"{cell_ref}: function {token.value} is not supported.",
                        UnsupportedWarning,
                        stacklevel=2,
                    )
                    return None

                ast_node.append(
                    {
                        "AST_node_type": "FUNCTION_NODE",
                        "AST_function_node_index": FUNCTION_NAME_TO_ID[token.value],
                        "AST_function_node_numArgs": token.num_args,
                    },
                )
            elif token.type == Token.OPERAND:
                func = getattr(formula, OPERAND_ARCHIVE_MAP[token.subtype])
                ast_node.append(func(token))

            elif token.type == Token.OP_IN:
                ast_node.append({"AST_node_type": OPERATOR_INFIX_MAP[token.value]})

        return model._formulas.lookup_key(
            table_id,
            TSCEArchives.FormulaArchive(**formula_attrs),
        )

    def add_table_xref_info(self, ref: dict[str, CellRange], node: dict) -> None:
        if not ref.name_scope_2:
            return

        sheet_name = (
            ref.name_scope_1
            if ref.name_scope_1
            else self._model.sheet_name(self._model.table_id_to_sheet_id(self._table_id))
        )
        table_uuid = self._model.table_name_to_uuid(sheet_name, ref.name_scope_2)
        xref_archive = NumbersUUID(table_uuid).protobuf4
        node["AST_cross_table_reference_extra_info"] = (
            TSCEArchives.ASTNodeArrayArchive.ASTCrossTableReferenceExtraInfoArchive(
                table_id=xref_archive,
            )
        )

    @staticmethod
    def _ast_sticky_bits(ref: dict[str, CellRange]) -> dict[str, str]:
        return {
            "begin_row_is_absolute": ref.row_start_is_abs,
            "begin_column_is_absolute": ref.col_start_is_abs,
            "end_row_is_absolute": ref.row_end_is_abs,
            "end_column_is_absolute": ref.col_end_is_abs,
        }

    def range_archive(self, token: "Token") -> dict:
        ref = parse_numbers_range(self._model, token.value)

        if ref.range_type == CellRangeType.RANGE:
            ast_colon_tract = {
                "preserve_rectangular": True,
                "relative_row": [{}],
                "relative_column": [{}],
                "absolute_row": [{}],
                "absolute_column": [{}],
            }

            if not (ref.col_start_is_abs and ref.col_end_is_abs):
                ast_colon_tract["relative_column"][0]["range_begin"] = (
                    (ref.col_end - self.col) if ref.col_start_is_abs else (ref.col_start - self.col)
                )

            if not (ref.col_start_is_abs) and not (ref.col_end_is_abs):
                ast_colon_tract["relative_column"][0]["range_end"] = ref.col_end - self.col

            if not (ref.row_start_is_abs and ref.row_end_is_abs):
                ast_colon_tract["relative_row"][0]["range_begin"] = ref.row_start - self.row
                if ref.row_start != ref.row_end:
                    ast_colon_tract["relative_row"][0]["range_end"] = ref.row_end - self.row

            if ref.col_start_is_abs or ref.col_end_is_abs:
                ast_colon_tract["absolute_column"][0]["range_begin"] = (
                    ref.col_start if ref.row_start_is_abs else ref.col_end
                )

            if ref.col_start_is_abs and ref.col_end_is_abs:
                ast_colon_tract["absolute_column"][0]["range_end"] = ref.col_end

            if ref.row_start_is_abs or ref.row_end_is_abs:
                ast_colon_tract["absolute_row"][0]["range_begin"] = (
                    ref.row_start if ref.row_start_is_abs else ref.row_end_is_abs
                )

            if ref.row_start_is_abs and ref.row_end_is_abs:
                ast_colon_tract["absolute_row"][0]["range_end"] = ref.row_end

            node = {
                "AST_node_type": "COLON_TRACT_NODE",
                "AST_sticky_bits": Formula._ast_sticky_bits(ref),
                "AST_colon_tract": ast_colon_tract,
            }

            key = (
                ref.col_start_is_abs,
                ref.col_end_is_abs,
                ref.row_start_is_abs,
                ref.row_end_is_abs,
            )
            ast_frozen_sticky_bits = {}
            if FROZEN_STICKY_BIT_MAP[key] is not None:
                sticky_bits = FROZEN_STICKY_BIT_MAP[key]
                ast_frozen_sticky_bits["begin_column_is_absolute"] = sticky_bits[0]
                ast_frozen_sticky_bits["end_column_is_absolute"] = sticky_bits[1]
                ast_frozen_sticky_bits["begin_row_is_absolute"] = sticky_bits[2]
                ast_frozen_sticky_bits["end_row_is_absolute"] = sticky_bits[3]
                node["AST_frozen_sticky_bits"] = ast_frozen_sticky_bits

            for key in ["absolute_row", "relative_row", "absolute_column", "relative_column"]:
                if len(ast_colon_tract[key][0].keys()) == 0:
                    del ast_colon_tract[key]

            self.add_table_xref_info(ref, node)

            return node

        if ref.range_type == CellRangeType.ROW_RANGE:
            row_start = ref.row_start if ref.row_start_is_abs else ref.row_start - self.row
            row_end = ref.row_end if ref.row_end_is_abs else ref.row_end - self.row

            node = {
                "AST_node_type": "COLON_TRACT_NODE",
                "AST_sticky_bits": Formula._ast_sticky_bits(ref),
                "AST_colon_tract": {
                    "relative_row": [{"range_begin": row_start, "range_end": row_end}],
                    "absolute_column": [{"range_begin": 0x7FFF}],
                    "preserve_rectangular": True,
                },
            }
            self.add_table_xref_info(ref, node)
            return node

        if ref.range_type == CellRangeType.COL_RANGE:
            col_start = ref.col_start if ref.col_start_is_abs else ref.col_start - self.col
            col_end = ref.col_end if ref.col_end_is_abs else ref.col_end - self.col

            node = {
                "AST_node_type": "COLON_TRACT_NODE",
                "AST_sticky_bits": Formula._ast_sticky_bits(ref),
                "AST_colon_tract": {
                    "relative_column": [{"range_begin": col_start, "range_end": col_end}],
                    "absolute_row": [{"range_begin": 2147483647}],
                    "preserve_rectangular": True,
                },
            }
            self.add_table_xref_info(ref, node)
            return node

        if ref.range_type == CellRangeType.NAMED_RANGE:
            _ = self._model.name_ref_cache.lookup_named_ref(self._table_id, ref)

            return {}

        if ref.range_type == CellRangeType.NAMED_ROW_COLUMN:
            _ = self._model.name_ref_cache.lookup_named_ref(self._table_id, ref)

            return {}

        # CellRangeType.CELL
        return {
            "AST_node_type": "CELL_REFERENCE_NODE",
            "AST_row": {
                "row": ref.row_start if ref.row_start_is_abs else ref.row_start - self.row,
                "absolute": ref.row_start_is_abs,
            },
            "AST_column": {
                "column": ref.col_start if ref.col_start_is_abs else ref.col_start - self.col,
                "absolute": ref.col_start_is_abs,
            },
        }

    def number_archive(self, token: "Token") -> ASTNodeArrayArchive.ASTNodeArchive:
        if float(token.value).is_integer():
            return {
                "AST_node_type": "NUMBER_NODE",
                "AST_number_node_number": int(float(token.value)),
                "AST_number_node_decimal_low": int(float(token.value)),
                "AST_number_node_decimal_high": 0x3040000000000000,
            }

        value = token.value
        exponent = (
            math.floor(math.log10(math.e) * math.log(abs(float(value))))
            if float(value) != 0.0
            else 0
        )
        if "E" in value:
            significand, exponent = value.split("E")
        else:
            significand = value
            exponent = 0
        num_dp = len(re.sub(r"0*$", "", str(significand).split(".")[1]))
        exponent = int(exponent) - num_dp
        decimal_low = int(float(significand) * 10**num_dp)
        decimal_high = ((DECIMAL128_BIAS * 2) + (2 * exponent)) << 48

        return {
            "AST_node_type": "NUMBER_NODE",
            "AST_number_node_number": float(value),
            "AST_number_node_decimal_low": decimal_low,
            "AST_number_node_decimal_high": decimal_high,
        }

    def text_archive(self, token: "Token") -> ASTNodeArrayArchive.ASTNodeArchive:
        # String literals from tokenizer include start and end quotes
        value = token.value[1:-1]
        # Numbers does not escape quotes in the AST
        value = value.replace('""', '"')
        return {
            "AST_node_type": "STRING_NODE",
            "AST_string_node_string": value,
        }

    def logical_archive(self, token: "Token") -> ASTNodeArrayArchive.ASTNodeArchive:
        if token.subtype == Token.LOGICAL:
            return {
                "AST_node_type": "BOOLEAN_NODE",
                "AST_boolean_node_boolean": token.value.lower() == "true",
            }

        return None

    def error(self, token: "Token") -> ASTNodeArrayArchive.ASTNodeArchive:
        return {
            "AST_node_type": "BOOLEAN_NODE",
            "AST_boolean_node_boolean": token.value.lower() == "true",
        }

    @staticmethod
    def formula_tokens(formula_str: str):
        tok = Tokenizer(formula_str.translate(OPERATOR_MAP))
        return Formula.rpn_tokens(tok.items)

    @staticmethod
    def rpn_tokens(tokens):
        output = []
        operators = []

        for token in tokens:
            if token.type in ["OPERAND", "NUMBER", "LITERAL", "TEXT", "RANGE"]:
                output.append(token)
                if operators and operators[-1].type == "FUNC":
                    operators[-1].num_args += 1
            elif token.type == "FUNC" and token.subtype == "OPEN":
                token.value = token.value[0:-1]
                operators.append(token)
                operators[-1].num_args = 0
            elif token.type in ["OPERATOR-POSTFIX", "OPERATOR-PREFIX"]:
                output.append(token)
            elif token.type == "OPERATOR-INFIX":
                while (
                    operators
                    and operators[-1].type == "OPERATOR-INFIX"
                    and OPERATOR_PRECEDENCE[operators[-1].value] >= OPERATOR_PRECEDENCE[token.value]
                ):
                    output.append(operators.pop())
                operators.append(token)
            elif token.type == "FUNC" and token.subtype == "CLOSE":
                while operators and (
                    operators[-1].type != "FUNC" and operators[-1].subtype != "OPEN"
                ):
                    output.append(operators.pop())
                if operators:
                    output.append(operators.pop())
            elif token.type == "SEP":
                if operators and operators[-1].type != "FUNC":
                    output.append(operators.pop())
            elif token.type == "PAREN":
                if token.subtype == "OPEN":
                    operators.append(token)
                elif token.subtype == "CLOSE":
                    while operators and operators[-1].subtype != "OPEN":
                        output.append(operators.pop())
                    operators.pop()
                    if operators and operators[-1].type == "FUNC":
                        output.append(operators.pop())

        while operators:
            output.append(operators.pop())

        return output

    def __str__(self) -> str:
        return "".join(reversed([str(x) for x in self._stack]))

    def pop(self) -> str:
        return self._stack.pop()

    def popn(self, num_args: int) -> tuple:
        values = ()
        for _ in range(num_args):
            values += (self._stack.pop(),)
        return values

    def push(self, val: str) -> None:
        self._stack.append(val)

    def add(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}+{arg2}")

    def array(self, *args) -> None:
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
            for _row_num in range(num_rows):
                args = self.popn(num_cols)
                args = ",".join(reversed(args))
                rows.append(f"{args}")
            args = ";".join(reversed(rows))
            self.push(f"{{{args}}}")

    def boolean(self, *args) -> None:
        node = args[2]
        if node.HasField("AST_token_node_boolean"):
            self.push(str(node.AST_token_node_boolean).upper())
        else:
            self.push(str(node.AST_boolean_node_boolean).upper())

    def concat(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}&{arg2}")

    def date(self, *args) -> None:
        # Date literals exported as DATE()
        node = args[2]
        dt = datetime(2001, 1, 1) + timedelta(seconds=node.AST_date_node_dateNum)  # noqa: DTZ001
        self.push(f"DATE({dt.year},{dt.month},{dt.day})")

    def div(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}÷{arg2}")

    def empty(self, *args) -> None:
        self.push("")

    def equals(self, *args) -> None:
        # Arguments appear to be reversed
        arg1, arg2 = self.popn(2)
        self.push(f"{arg2}={arg1}")

    def function(self, *args) -> None:
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

    def greater_than(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}>{arg2}")

    def greater_than_or_equal(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≥{arg2}")

    def less_than(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}<{arg2}")

    def less_than_or_equal(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≤{arg2}")

    def list(self, *args) -> None:
        node = args[2]
        args = self.popn(node.AST_list_node_numArgs)
        args = ",".join(reversed([str(x) for x in args]))
        self.push(f"({args})")

    def mul(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}×{arg2}")

    def negate(self, *args) -> None:
        arg1 = self.pop()
        self.push(f"-{arg1}")

    def not_equals(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}≠{arg2}")

    def number(self, *args) -> None:
        node = args[2]
        if node.AST_number_node_decimal_high == 0x3040000000000000:
            # Integer: don't use decimals
            self.push(str(node.AST_number_node_decimal_low))
        else:
            self.push(number_to_str(node.AST_number_node_number))

    def percent(self, *args) -> None:
        arg1 = self.pop()
        self.push(f"{arg1}%")

    def power(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}^{arg2}")

    def range(self, *args) -> None:
        arg2, arg1 = [str(x) for x in self.popn(2)]
        func_range = "(" in arg1 or "(" in arg2
        if "::" in arg1 and not func_range:
            # Assumes references are not cross-table
            arg1_parts = arg1.split("::")
            arg2_parts = arg2.split("::")
            self.push(f"{arg1_parts[0]}::{arg1_parts[1]}:{arg2_parts[1]}")
        else:
            self.push(f"{arg1}:{arg2}")

    def string(self, *args) -> None:
        node = args[2]
        # Numbers does not escape quotes in the AST; in the app, they are
        # doubled up just like in Excel
        value = node.AST_string_node_string.replace('"', '""')
        self.push(f'"{value}"')

    def sub(self, *args) -> None:
        arg2, arg1 = self.popn(2)
        self.push(f"{arg1}-{arg2}")

    def xref(self, *args) -> None:
        (row, col, node) = args
        self.push(self._model.node_to_ref(self._table_id, row, col, node))


NODE_FUNCTION_MAP = {
    "ADDITION_NODE": "add",
    "APPEND_WHITESPACE_NODE": None,
    "ARRAY_NODE": "array",
    "BEGIN_EMBEDDED_NODE_ARRAY": None,
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
    "PREPEND_WHITESPACE_NODE": None,
    "STRING_NODE": "string",
    "SUBTRACTION_NODE": "sub",
    "TOKEN_NODE": "boolean",
}


class TableFormulas:
    def __init__(self, model, table_id) -> None:
        self._model = model
        self._table_id = table_id
        self._formula_type_lookup = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

    def formula(self, formula_key, row, col):
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
        number = re.sub(r"[,-.]", "", number)
        zeroes = "0" * (abs(int(exp)) - 1)
        if int(exp) > 0:
            return f"{number}{zeroes}"
        return f"0.{zeroes}{number}"
    return v_str
