import re
import warnings
from datetime import datetime, timedelta

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.generated.TSCEArchives_pb2 import ASTNodeArrayArchive
from numbers_parser.numbers_uuid import NumbersUUID
from numbers_parser.tokenizer import Token, Tokenizer, parse_numbers_range

FUNCTION_NAME_TO_ID = {v: k for k, v in FUNCTION_MAP.items()}

OPERATOR_MAP = str.maketrans({"×": "*", "÷": "/", "≥": ">=", "≤": "<=", "≠": "<>"})

OPERATOR_PRECEDENCE = {"%": 6, "^": 5, "×": 4, "*": 4, "/": 4, "÷": 4, "+": 3, "-": 3, "&": 2}

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
}


class Formula(list):
    def __init__(self, model, table_id, row, col) -> None:
        self._stack = []
        self._model = model
        self._table_id = table_id
        self.row = row
        self.col = col

    @classmethod
    def from_str(cls, model, table_id, row, col, formula_str) -> "Formula":
        formula = cls(model, table_id, row, col)
        formula._tokens = cls.formula_tokens(formula_str)
        archive = TSCEArchives.FormulaArchive()
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
                archive.AST_node_array.AST_node.append(
                    ASTNodeArrayArchive.ASTNodeArchive(
                        AST_node_type="FUNCTION_NODE",
                        AST_function_node_index=FUNCTION_NAME_TO_ID[token.value],
                        AST_function_node_numArgs=token.num_args,
                    ),
                )
            elif token.type == Token.OPERAND:
                method = OPERAND_ARCHIVE_MAP[token.subtype]
                node = getattr(cls, method)(model, table_id, row, col, token)
                # node = Formula.operand_archive(row, col, token)
                # # if node is not None:
                archive.AST_node_array.AST_node.append(node)
            elif token.type == Token.OP_IN:
                archive.AST_node_array.AST_node.append(
                    ASTNodeArrayArchive.ASTNodeArchive(
                        AST_node_type=OPERATOR_INFIX_MAP[token.value],
                    ),
                )
        formula._archive = archive
        return formula

    @staticmethod
    def range_archive(
        model: object,
        table_id: int,
        row: int,
        col: int,
        token: "Token",
    ) -> ASTNodeArrayArchive.ASTNodeArchive:
        r = parse_numbers_range(token.value)
        if r["range_end"]:
            row_range_begin = r["row_start"] if r["row_start_abs"] else r["row_start"] - row
            row_range_end = r["row_end"] if r["row_end_abs"] else r["row_end"] - row
            col_range_begin = r["col_start"] if r["col_start_abs"] else r["col_start"] - col
            col_range_end = r["col_end"] if r["col_end_abs"] else r["col_end"] - col

            args = {}
            if r["named_ref1"]:
                if r["named_ref2"] is None:
                    sheet_id = model.table_id_to_sheet_id(table_id)
                    sheet_name = model.sheet_name(sheet_id)
                    table_name = r["named_ref1"]
                    table_uuid = model.table_name_to_uuid(sheet_name, table_name)
                else:
                    table_uuid = model.table_name_to_uuid(r["named_ref1"], r["named_ref2"])
                xref_archive = NumbersUUID(table_uuid).protobuf4
                args["AST_cross_table_reference_extra_info"] = (
                    TSCEArchives.ASTNodeArrayArchive.ASTCrossTableReferenceExtraInfoArchive(
                        table_id=xref_archive,
                    )
                )

            row_range = {"range_begin": row_range_begin}
            if row_range_begin != row_range_end:
                row_range["range_end"] = row_range_end

            return ASTNodeArrayArchive.ASTNodeArchive(
                AST_node_type="COLON_TRACT_NODE",
                AST_sticky_bits=ASTNodeArrayArchive.ASTStickyBits(
                    begin_row_is_absolute=r["row_start_abs"],
                    begin_column_is_absolute=r["col_start_abs"],
                    end_row_is_absolute=r["row_end_abs"],
                    end_column_is_absolute=r["col_end_abs"],
                ),
                AST_colon_tract=ASTNodeArrayArchive.ASTColonTractArchive(
                    relative_row=[
                        ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractRelativeRangeArchive(
                            **row_range,
                        ),
                    ],
                    relative_column=[
                        ASTNodeArrayArchive.ASTColonTractArchive.ASTColonTractRelativeRangeArchive(
                            range_begin=col_range_begin,
                            range_end=col_range_end,
                        ),
                    ],
                    preserve_rectangular=True,
                ),
                **args,
            )
        return ASTNodeArrayArchive.ASTNodeArchive(
            AST_node_type="CELL_REFERENCE_NODE",
            AST_row=ASTNodeArrayArchive.ASTRowCoordinateArchive(
                row=r["row_start"] if r["row_start_abs"] else r["row_start"] - row,
                absolute=r["row_start_abs"],
            ),
            AST_column=ASTNodeArrayArchive.ASTColumnCoordinateArchive(
                column=r["col_start"] if r["col_start_abs"] else r["col_start"] - col,
                absolute=r["col_start_abs"],
            ),
        )

    @staticmethod
    def number_archive(
        _model: object,
        _table_id: int,
        _row: int,
        _col: int,
        token: "Token",
    ) -> ASTNodeArrayArchive.ASTNodeArchive:
        if float(token.value).is_integer():
            return ASTNodeArrayArchive.ASTNodeArchive(
                AST_node_type="NUMBER_NODE",
                AST_number_node_number=int(token.value),
                AST_number_node_decimal_low=int(token.value),
                AST_number_node_decimal_high=0x3040000000000000,
            )
        return ASTNodeArrayArchive.ASTNodeArchive(
            AST_node_type="NUMBER_NODE",
            AST_number_node_number=float(token.value),
            # AST_number_node_decimal_low=int(token.value),
            # AST_number_node_decimal_high=0x3040000000000000,
        )

    @staticmethod
    def text_archive(
        _model: object,
        _table_id: int,
        _row: int,
        _col: int,
        token: "Token",
    ) -> ASTNodeArrayArchive.ASTNodeArchive:
        # String literals from tokenizer include start and end quotes
        value = token.value[1:-1]
        # Numbers does not escape quotes in the AST
        value = value.replace('""', '"')
        return ASTNodeArrayArchive.ASTNodeArchive(
            AST_node_type="STRING_NODE",
            AST_string_node_string=value,
        )

    @staticmethod
    def logical_archive(
        _model: object,
        _table_id: int,
        _row: int,
        _col: int,
        token: "Token",
    ) -> ASTNodeArrayArchive.ASTNodeArchive:
        if token.subtype == Token.LOGICAL:
            return ASTNodeArrayArchive.ASTNodeArchive(
                AST_node_type="BOOLEAN_NODE",
                AST_boolean_node_boolean=token.value.lower() == "true",
            )

        return None

    @staticmethod
    def formula_tokens(formula_str: str):
        formula_str = formula_str.translate(OPERATOR_MAP)
        tok = Tokenizer(formula_str if formula_str.startswith("=") else "=" + formula_str)
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
        return "".join(reversed(self._stack))

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
        args = ",".join(reversed(args))
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
        args = ",".join(reversed(args))
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
        arg2, arg1 = self.popn(2)
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
    # Unimplemented: AVERAGE
    # Unimplemented: AVERAGE_ALL
    "BEGIN_EMBEDDED_NODE_ARRAY": None,
    # Unimplemented: BODY_ROWS
    "BOOLEAN_NODE": "boolean",
    # Unimplemented: CATEGORY_REF_NODE
    "CELL_REFERENCE_NODE": "xref",
    # Unimplemented: CHART_GROUP_VALUE_HIERARCHY
    "COLON_NODE": "range",
    "COLON_NODE_WITH_UIDS": "range",
    "COLON_TRACT_NODE": "xref",
    # Unimplemented: COLON_WITH_UIDS_NODE
    "CONCATENATION_NODE": "concat",
    # Unimplemented: COUNT_ALL
    # Unimplemented: COUNT_BLANK
    # Unimplemented: COUNT_DUPS
    # Unimplemented: COUNT_NO_ALL
    # Unimplemented: COUNT_ROWS
    # Unimplemented: COUNT_UNIQUE
    # Unimplemented: CROSS_TABLE_CELL_REFERENCE_NODE
    "DATE_NODE": "date",
    "DIVISION_NODE": "div",
    # Unimplemented: DURATION_NODE
    "EMPTY_ARGUMENT_NODE": "empty",
    "END_THUNK_NODE": None,
    "EQUAL_TO_NODE": "equals",
    "FUNCTION_NODE": "function",
    "GREATER_THAN_NODE": "greater_than",
    "GREATER_THAN_OR_EQUAL_TO_NODE": "greater_than_or_equal",
    # Unimplemented: GROUP_VALUE
    # Unimplemented: GROUP_VALUE_HIERARCHY
    # Unimplemented: GSCE.CalculationEngineAstNodeType={ADDITION_NODE
    # Unimplemented: INDIRECT
    # Unimplemented: LABEL
    "LESS_THAN_NODE": "less_than",
    "LESS_THAN_OR_EQUAL_TO_NODE": "less_than_or_equal",
    # Unimplemented: LINKED_CELL_REF_NODE
    # Unimplemented: LINKED_COLUMN_REF_NODE
    # Unimplemented: LINKED_ROW_REF_NODE
    "LIST_NODE": "list",
    # Unimplemented: LOCAL_CELL_REFERENCE_NODE
    # Unimplemented: MAX
    # Unimplemented: MEDIAN
    # Unimplemented: MIN
    # Unimplemented: MISSING_RUNNING_TOTAL_IN_FIELD
    # Unimplemented: MODE
    "MULTIPLICATION_NODE": "mul",
    "NEGATION_NODE": "negate",
    # Unimplemented: NONE
    "NOT_EQUAL_TO_NODE": "not_equals",
    "NUMBER_NODE": "number",
    "PERCENT_NODE": "percent",
    # Unimplemented: PLUS_SIGN_NODE
    "POWER_NODE": "power",
    "PREPEND_WHITESPACE_NODE": None,
    # Unimplemented: PRODUCT
    # Unimplemented: RANGE
    # Unimplemented: REFERENCE_ERROR_NODE
    # Unimplemented: REFERENCE_ERROR_WITH_UIDS_NODE
    "STRING_NODE": "string",
    # Unimplemented: ST_DEV
    # Unimplemented: ST_DEV_ALL
    # Unimplemented: ST_DEV_POP
    # Unimplemented: ST_DEV_POP_ALL
    "SUBTRACTION_NODE": "sub",
    # Unimplemented: THUNK_NODE
    "TOKEN_NODE": "boolean",
    # Unimplemented: TOTAL
    # Unimplemented: UID_REFERENCE_NODE
    # Unimplemented: UNKNOWN_FUNCTION_NODE
    # Unimplemented: VARIANCE
    # Unimplemented: VARIANCE_ALL
    # Unimplemented: VARIANCE_POP
    # Unimplemented: VARIANCE_POP_ALL
    # Unimplemented: VIEW_TRACT_REF_NODE
}


class TableFormulas:
    def __init__(self, model, table_id) -> None:
        self._model = model
        self._table_id = table_id
        self._formula_type_lookup = {
            k: v.name
            for k, v in TSCEArchives._ASTNODEARRAYARCHIVE_ASTNODETYPE.values_by_number.items()
        }

    def is_formula(self, row, col):
        return (row, col) in self._model.formula_cell_ranges(self._table_id)

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


def str_to_number(v: str) -> int:
    """Convert a string to a number."""
    if "." in v:
        v = v.split(".")
        return int(v[0] + v[1])
    return int(v)
