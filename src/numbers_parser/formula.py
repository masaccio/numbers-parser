import re
import warnings
from datetime import datetime, timedelta

from numbers_parser.exceptions import UnsupportedWarning
from numbers_parser.generated import TSCEArchives_pb2 as TSCEArchives
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.generated.TSCEArchives_pb2 import ASTNodeArrayArchive

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
                node = Formula.operand_archive(row, col, token)
                # if node is not None:
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
    def operand_archive(row: int, col: int, token: "Token") -> ASTNodeArrayArchive.ASTNodeArchive:
        if token.subtype == Token.RANGE:
            r = parse_numbers_range(token.value)
            if r["range_end"]:
                row_range_begin = r["row_start"] if r["row_start_abs"] else r["row_start"] - row
                col_range_begin = r["col_start"] if r["col_start_abs"] else r["col_start"] - col
                col_range_end = r["col_end"] if r["col_end_abs"] else r["col_end"] - col
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
                                range_begin=row_range_begin,
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
        if token.subtype == Token.NUMBER:
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
        if token.subtype == Token.TEXT:
            # String literals from tokenizer include start and end quotes
            value = token.value[1:-1]
            # Numbers does not escape quotes in the AST
            value = value.replace('""', '"')
            return ASTNodeArrayArchive.ASTNodeArchive(
                AST_node_type="STRING_NODE",
                AST_string_node_string=value,
            )

        if token.subtype == Token.LOGICAL:
            return ASTNodeArrayArchive.ASTNodeArchive(
                AST_node_type="BOOLEAN_NODE",
                AST_boolean_node_boolean=token.value.lower() == "true",
            )

        return None

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

    @staticmethod
    def formula_tokens(formula_str: str):
        formula_str = formula_str.translate(OPERATOR_MAP)
        tok = Tokenizer(formula_str if formula_str.startswith("=") else "=" + formula_str)
        return Formula.rpn_tokens(tok.items)

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


# The Tokenizer and Token classesare taken from the openpyxl library which is
# licensed under the MIT License. The original source code can be found at:
#
# https://github.com/gleeda/openpyxl/blob/master/openpyxl/formula/tokenizer.py
#
# Copyright (c) 2010 openpyxl
#
# The openpyxl tokenizer is based on the Javascript tokenizer originally found at
# http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html written by Eric
# Bachtal, and now archived by the Internet Archive at the following URL:
#
# https://archive.is/OCsys


class TokenizerError(Exception):
    """Base class for all Tokenizer errors."""


class Tokenizer:
    """
    A tokenizer for Excel worksheet formulae.

    Converts a unicode string representing an Excel formula (in A1 notation)
    into a sequence of `Token` objects.

    `formula`: The unicode string to tokenize

    Tokenizer defines a method `.parse()` to parse the formula into tokens,
    which can then be accessed through the `.items` attribute.

    """

    SN_RE = re.compile("^[1-9](\\.[0-9]+)?E$")  # Scientific notation
    WSPACE_RE = re.compile(" +")
    STRING_REGEXES = {
        # Inside a string, all characters are treated as literals, except for
        # the quote character used to start the string. That character, when
        # doubled is treated as a single character in the string. If an
        # unmatched quote appears, the string is terminated.
        '"': re.compile('"(?:[^"]*"")*[^"]*"(?!")'),
        "'": re.compile("'(?:[^']*'')*[^']*'(?!')"),
    }
    ERROR_CODES = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")
    TOKEN_ENDERS = ",;})+-*/^&=><%×÷≥≤≠"  # Each of these characters, marks the  # noqa: S105
    # end of an operand token

    def __init__(self, formula):
        self.formula = formula
        self.items = []
        self.token_stack = []  # Used to keep track of arrays, functions, and
        # parentheses
        self.offset = 0  # How many chars have we read
        self.token = []  # Used to build up token values char by char
        self.parse()

    def __repr__(self):
        item_str = ",".join([repr(token) for token in self.items])
        return f"[{item_str}]"

    def parse(self):
        """Populate self.items with the tokens from the formula."""
        if not self.formula:
            return
        if self.formula[0] == "=":
            self.offset += 1
        else:
            self.items.append(Token(self.formula, Token.LITERAL))
            return
        consumers = (
            ("\"'", self.parse_string),
            ("[", self.parse_brackets),
            ("#", self.parse_error),
            # (" ", self.parse_whitespace),
            ("+-*/^&=><%×÷≥≤≠", self.parse_operator),
            ("{(", self.parse_opener),
            (")}", self.parse_closer),
            (";,", self.parse_separator),
        )
        dispatcher = {}  # maps chars to the specific parsing function
        for chars, consumer in consumers:
            dispatcher.update(dict.fromkeys(chars, consumer))
        while self.offset < len(self.formula):
            if self.check_scientific_notation():  # May consume one character
                continue
            curr_char = self.formula[self.offset]
            if curr_char in self.TOKEN_ENDERS:
                self.save_token()
            if curr_char in dispatcher:
                self.offset += dispatcher[curr_char]()
            else:
                # TODO: this can probably be sped up using a regex to get to
                # the next interesting character
                self.token.append(curr_char)
                self.offset += 1
        self.save_token()

    def parse_string(self):
        """
        Parse a "-delimited string or '-delimited link.

        The offset must be pointing to either a single quote ("'") or double
        quote ('"') character. The strings are parsed according to Excel
        rules where to escape the delimiter you just double it up. E.g.,
        "abc""def" in Excel is parsed as 'abc"def' in Python.

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        self.assert_empty_token()
        delim = self.formula[self.offset]
        if delim not in ('"', "'"):
            msg = f"Invalid string delimiter: {delim}"
            raise TokenizerError(msg)
        regex = self.STRING_REGEXES[delim]
        match = regex.match(self.formula[self.offset :])
        if match is None:
            subtype = "string" if delim == '"' else "link"
            msg = f"Reached end of formula while parsing {subtype} in {self.formula}"
            raise TokenizerError(msg)
        match = match.group(0)
        if delim == '"':
            self.items.append(Token.make_operand(match))
        else:
            self.token.append(match)
        return len(match)

    def parse_brackets(self):
        """
        Consume all the text between square brackets [].

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        if self.formula[self.offset] != "[":
            msg = f"Expected '[', found: {self.formula[self.offset]}"
            raise TokenizerError(msg)
        right = self.formula.find("]", self.offset) + 1
        if right == 0:
            msg = "Encountered unmatched '[' in '{self.formula}'"
            raise TokenizerError(msg)
        self.token.append(self.formula[self.offset : right])
        return right - self.offset

    def parse_error(self):
        """
        Consume the text following a '#' as an error.

        Looks for a match in self.ERROR_CODES and returns the number of
        characters matched. (Does not update self.offset)

        """
        self.assert_empty_token()
        if self.formula[self.offset] != "#":
            msg = f"Expected '#', found: {self.formula[self.offset]}"
            raise TokenizerError(msg)
        subformula = self.formula[self.offset :]
        for err in self.ERROR_CODES:
            if subformula.startswith(err):
                self.items.append(Token.make_operand(err))
                return len(err)
        msg = f"Invalid error code at position {self.offset} in 'self.formula'"
        raise TokenizerError(msg)

    def parse_whitespace(self):
        """
        Consume a string of consecutive spaces.

        Returns the number of spaces found. (Does not update
        self.offset).

        """
        if self.formula[self.offset] != " ":
            msg = f"Expected ' ', found: {self.formula[self.offset]}"
            raise TokenizerError(msg)
        self.items.append(Token(" ", Token.WSPACE))
        return self.WSPACE_RE.match(self.formula[self.offset :]).end()

    def parse_operator(self):
        """
        Consume the characters constituting an operator.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        if self.formula[self.offset : self.offset + 2] in (">=", "<=", "<>", "≥", "≤", "≠"):
            self.items.append(
                Token(
                    self.formula[self.offset : self.offset + 2],
                    Token.OP_IN,
                ),
            )
            return 2
        curr_char = self.formula[self.offset]  # guaranteed to be 1 char
        if curr_char == "%":
            token = Token("%", Token.OP_POST)
        elif curr_char in "*/^&=><×÷≥≤≠":
            token = Token(curr_char, Token.OP_IN)
        # From here on, curr_char is guaranteed to be in '+-'
        elif not self.items:
            token = Token(curr_char, Token.OP_PRE)
        else:
            prev = self.items[-1]
            is_infix = prev.subtype == Token.CLOSE or prev.type in (Token.OP_POST, Token.OPERAND)
            token = Token(curr_char, Token.OP_IN) if is_infix else Token(curr_char, Token.OP_PRE)
        self.items.append(token)
        return 1

    def parse_opener(self):
        """
        Consumes a ( or { character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        if self.formula[self.offset] not in ("(", "{"):
            msg = f"Expected '(' or '{{', found: {self.formula[self.offset]}"
            raise TokenizerError(msg)
        if self.formula[self.offset] == "{":
            self.assert_empty_token()
            token = Token.make_subexp("{")
        elif self.token:
            token_value = "".join(self.token) + "("
            del self.token[:]
            token = Token.make_subexp(token_value)
        else:
            token = Token.make_subexp("(")
        self.items.append(token)
        self.token_stack.append(token)
        return 1

    def parse_closer(self):
        """
        Consumes a } or ) character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        if self.formula[self.offset] not in (")", "}"):
            msg = f"Expected ')' or '}}', found: {self.formula[self.offset]}"
            raise TokenizerError(msg)
        token = self.token_stack.pop().get_closer()
        if token.value != self.formula[self.offset]:
            msg = "Mismatched ( and { pair in '{self.formula}'"
            raise TokenizerError(msg)
        self.items.append(token)
        return 1

    def parse_separator(self):
        """
        Consumes a ; or , character.

        Returns the number of charactes consumed. (Does not update
        self.offset)

        """
        curr_char = self.formula[self.offset]
        if curr_char not in (";", ","):
            msg = f"Expected ';' or ',', found: {curr_char}"
            raise TokenizerError(msg)
        if curr_char == ";":
            token = Token.make_separator(";")
        else:
            try:
                top_type = self.token_stack[-1].type
            except IndexError:
                token = Token(",", Token.OP_IN)  # Range Union operator
            else:
                if top_type == Token.PAREN:
                    token = Token(",", Token.OP_IN)  # Range Union operator
                else:
                    token = Token.make_separator(",")
        self.items.append(token)
        return 1

    def check_scientific_notation(self):
        """
        Consumes a + or - character if part of a number in sci. notation.

        Returns True if the character was consumed and self.offset was
        updated, False otherwise.

        """
        curr_char = self.formula[self.offset]
        if curr_char in "+-" and len(self.token) >= 1 and self.SN_RE.match("".join(self.token)):
            self.token.append(curr_char)
            self.offset += 1
            return True
        return False

    def assert_empty_token(self):
        """
        Ensure that there's no token currently being parsed.

        If there are unconsumed token contents, it means we hit an unexpected
        token transition. In this case, we raise a TokenizerError

        """
        if self.token:
            msg = "Unexpected character at position {self.offset} in 'self.formula'"
            raise TokenizerError(msg)

    def save_token(self):
        """If there's a token being parsed, add it to the item list."""
        if self.token:
            self.items.append(Token.make_operand("".join(self.token)))
            del self.token[:]

    def render(self):
        """Convert the parsed tokens back to a string."""
        if not self.items:
            return ""
        if self.items[0].type == Token.LITERAL:
            return self.items[0].value
        return "=" + "".join(token.value for token in self.items)


class Token:
    """
    A token in an Excel formula.

    Tokens have three attributes:

    * `value`: The string value parsed that led to this token
    * `type`: A string identifying the type of token
    * `subtype`: A string identifying subtype of the token (optional, and
                 defaults to "")

    """

    __slots__ = ["num_args", "subtype", "type", "value"]

    LITERAL = "LITERAL"
    OPERAND = "OPERAND"
    FUNC = "FUNC"
    ARRAY = "ARRAY"
    PAREN = "PAREN"
    SEP = "SEP"
    OP_PRE = "OPERATOR-PREFIX"
    OP_IN = "OPERATOR-INFIX"
    OP_POST = "OPERATOR-POSTFIX"
    WSPACE = "WHITE-SPACE"

    def __init__(self, value, type_, subtype=""):
        self.value = value
        self.type = type_
        self.subtype = subtype
        self.num_args = 0

    def __repr__(self):
        return f"{self.type}({self.subtype},'{self.value}')"

    # Literal operands:
    #
    # Literal operands are always of type 'OPERAND' and can be of subtype
    # 'TEXT' (for text strings), 'NUMBER' (for all numeric types), 'LOGICAL'
    # (for TRUE and FALSE), 'ERROR' (for literal error values), or 'RANGE'
    # (for all range references).

    TEXT = "TEXT"
    NUMBER = "NUMBER"
    LOGICAL = "LOGICAL"
    ERROR = "ERROR"
    RANGE = "RANGE"

    @classmethod
    def make_operand(cls, value):
        """Create an operand token."""
        if value.startswith('"'):
            subtype = cls.TEXT
        elif value.startswith("#"):
            subtype = cls.ERROR
        elif value in ("TRUE", "FALSE"):
            subtype = cls.LOGICAL
        else:
            try:
                float(value)
                subtype = cls.NUMBER
            except ValueError:
                subtype = cls.RANGE
        return cls(value, cls.OPERAND, subtype)

    # Subexpresssions
    #
    # There are 3 types of `Subexpressions`: functions, array literals, and
    # parentheticals. Subexpressions have 'OPEN' and 'CLOSE' tokens. 'OPEN'
    # is used when parsing the initital expression token (i.e., '(' or '{')
    # and 'CLOSE' is used when parsing the closing expression token ('}' or
    # ')').

    OPEN = "OPEN"
    CLOSE = "CLOSE"

    @classmethod
    def make_subexp(cls, value, func=False):
        """
        Create a subexpression token.

        `value`: The value of the token
        `func`: If True, force the token to be of type FUNC

        """
        if value[-1] not in ("{", "}", "(", ")"):
            msg = f"Invalid subexpression value: {value}"
            raise TokenizerError(msg)
        if func:
            if not re.match(".+\\(|\\)", value):
                msg = f"Invalid function subexpression value: {value}"
                raise TokenizerError(msg)
            type_ = Token.FUNC
        elif value in "{}":
            type_ = Token.ARRAY
        elif value in "()":
            type_ = Token.PAREN
        else:
            type_ = Token.FUNC
        subtype = cls.CLOSE if value in ")}" else cls.OPEN
        return cls(value, type_, subtype)

    def get_closer(self):
        """Return a closing token that matches this token's type."""
        if self.type not in (self.FUNC, self.ARRAY, self.PAREN):
            msg = f"Invalid token type for closer: {self.type}"
            raise TokenizerError(msg)
        if self.subtype != self.OPEN:
            msg = f"Invalid token subtype for closer: {self.subtype}"
            raise TokenizerError(msg)
        value = "}" if self.type == self.ARRAY else ")"
        return self.make_subexp(value, func=self.type == self.FUNC)

    # Separator tokens
    #
    # Argument separators always have type 'SEP' and can have one of two
    # subtypes: 'ARG', 'ROW'. 'ARG' is used for the ',' token, when used to
    # delimit either function arguments or array elements. 'ROW' is used for
    # the ';' token, which is always used to delimit rows in an array
    # literal.

    ARG = "ARG"
    ROW = "ROW"

    @classmethod
    def make_separator(cls, value):
        """Create a separator token"""
        if value not in (",", ";"):
            msg = f"Invalid separator value: {value}"
            raise TokenizerError(msg)

        subtype = cls.ARG if value == "," else cls.ROW
        return cls(value, cls.SEP, subtype)


# Regex pattern
SHEET_RANGE_REGEXP = re.compile(
    r"""
    ^                                             # Start of the string
    (?:(?P<sheet_name>[^:]+)::)?                  # Optional sheet name followed by ::
    (?:(?P<table_name>[^:]+)::)?                  # Optional table name followed by ::
    (?:
        (?P<range_start>                          # Group for single cell or start of range
            (?P<col_start>\$?([0-9]+|[A-Z]+))     # Start column (e.g., A, $A, AA), $ optional
            (?P<row_start>\$?\d+)                 # Start row (e.g., 1, $1), $ optional
        )
        (?:
            :
            (?P<range_end>                        # Optional range end (e.g., B2)
                (?P<col_end>\$?([0-9]+|[A-Z]+))?  # End column (e.g., B, $B, AB), $ optional
                (?P<row_end>\$?\d+)?              # End row (e.g., 2, $2), $ optional
            )
        )?
    |
        (?P<named_range>[a-zA-Z_][a-zA-Z0-9_]*)   # Named range (e.g., cats)
    )
    $                                             # End of the string
""",
    re.VERBOSE,
)

# sheet_name: Matches the optional sheet name (e.g., "Sheet1").
# table_name: Matches the optional table name (e.g., "Table1").
# range_start: Matches the start of a range or single cell (e.g., "A1" or "$B$2").
# col_start: Matches the column portion of the range start (e.g., "A", "$B").
# row_start: Matches the row portion of the range start (e.g., "1", "$2").
# range_end: Matches the optional end of a range (e.g., "C3").
# col_end: Matches the column portion of the range end (e.g., "C", "$D").
# row_end: Matches the row portion of the range end (e.g., "3", "$4").
# named_range: Matches a named range (e.g., "cats").


def parse_numbers_range(range_str: str) -> dict:
    """
    Parse a cell range string in Numbers format.

    Args:
        range_str (str): The Numbers cell range string.

    Returns:
        dict: A dictionary containing the start and end column/row numbers
              with zero offset, booleans indicating whether the references
              are absolute or relative, and any sheet or table name.

    """
    if not (match := SHEET_RANGE_REGEXP.match(range_str)):
        msg = f"Invalid range string: {range_str}"
        raise ValueError(msg)

    def col_to_index(col: str) -> int:
        if col is None:
            return ""
        col = col.lstrip("$")
        index = 0
        for char in col:
            index = index * 26 + (ord(char) - ord("A") + 1)
        return index - 1

    def row_to_index(row: str) -> int:
        return int(row.lstrip("$")) - 1 if row is not None else ""

    col_start = match.group("col_start") or ""
    row_start = match.group("row_start") or ""
    col_end = match.group("col_end") or ""
    row_end = match.group("row_end") or ""

    return {
        **match.groupdict(),
        "col_start": col_to_index(col_start),
        "row_start": row_to_index(row_start),
        "col_end": col_to_index(col_end) if col_end else col_to_index(col_start),
        "end_row": row_to_index(row_end) if row_end else row_to_index(row_start),
        "col_start_abs": col_start.startswith("$"),
        "row_start_abs": row_start.startswith("$"),
        "col_end_abs": col_end.startswith("$") if col_end else col_start.startswith("$"),
        "row_end_abs": row_end.startswith("$") if row_end else row_start.startswith("$"),
    }
