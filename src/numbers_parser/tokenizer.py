import re

SHEET_RANGE_REGEXP = re.compile(
    r"""
    ^                                             # Start of the string
    (?:(?P<named_ref1>[^:]+)::)?                  # Optional sheet/table name followed by ::
    (?:(?P<named_ref2>[^:]+)::)?                  # Optional table name followed by ::
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

# named_ref1: Matches optional sheet name (e.g., "Sheet1"), or table if in current sheet
# named_ref2: Matches optional table name (e.g., "Table1").
# range_start: Matches start of a range or single cell (e.g., "A1" or "$B$2").
# col_start: Matches column portion of the range start (e.g., "A", "$B").
# row_start: Matches row portion of the range start (e.g., "1", "$2").
# range_end: Matches optional end of a range (e.g., "C3").
# col_end: Matches column portion of the range end (e.g., "C", "$D").
# row_end: Matches row portion of the range end (e.g., "3", "$4").
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
        "row_end": row_to_index(row_end) if row_end else row_to_index(row_start),
        "col_start_abs": col_start.startswith("$"),
        "row_start_abs": row_start.startswith("$"),
        "col_end_abs": col_end.startswith("$") if col_end else col_start.startswith("$"),
        "row_end_abs": row_end.startswith("$") if row_end else row_start.startswith("$"),
    }


# The Tokenizer and Token classes are taken from the openpyxl library which is
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
