import re
from sys import argv

from numbers_parser import Document

SHEET_RANGE_REGEXP = r"""
    (?:(?P<sheet_name>\w[^:]+)::)?                # Optional sheet name followed by ::
    (?:(?P<table_name>\w[^:]+)::)?                # Optional table name followed by ::
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
"""


def print_formula(formula: str) -> None:
    sheet_idx = 1
    table_idx = 1
    for match in re.finditer(SHEET_RANGE_REGEXP, formula, re.VERBOSE):
        if match.group("sheet_name"):
            formula = formula.replace(match.group("sheet_name"), f"XSHEET{sheet_idx}")
            sheet_idx += 1
        if match.group("table_name"):
            formula = formula.replace(match.group("table_name"), f"XTABLE{table_idx}")
            table_idx += 1
    formula = formula.replace('"""', "§")
    formula = re.sub(r'".*?"', '"XSTRING"', formula)
    formula = formula.replace("§", '"""')
    print(formula)


for filename in argv[1:]:
    doc = Document(filename)
    for sheet in doc.sheets:
        for table in sheet.tables:
            for row in table.rows():
                for cell in row:
                    if cell is not None and cell.formula:
                        print_formula(cell.formula)
