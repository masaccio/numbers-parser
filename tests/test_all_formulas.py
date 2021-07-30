import pytest
import pytest_check as check

from numbers_parser import Document
from numbers_parser.cell import ErrorCell, TextCell

DOCUMENT = "tests/data/test-all-forumulas.numbers"


def compare_table_functions(sheet_name):
    doc = Document(DOCUMENT)
    sheet = doc.sheets()[sheet_name]
    table = sheet.tables()["Tests"]
    for i, row in enumerate(table.rows()):
        if i == 0 or not row[3].value:
            # Skip header and invalid test rows
            continue

        formula_text = row[0].value
        formula_result = row[1].value
        formula_ref_value = row[2].value
        check.is_true(row[1].has_formula)
        check.equal(row[1].formula, formula_text)
        if isinstance(row[1], ErrorCell):
            pass
        elif isinstance(row[1], TextCell) and formula_ref_value is None:
            pass
        else:
            check.equal(formula_result, formula_ref_value)


# @pytest.mark.experimental
def test_information_functions():
    compare_table_functions("Information")


@pytest.mark.experimental
def test_text_functions():
    compare_table_functions("Text")


# @pytest.mark.experimental
def test_math_functions():
    compare_table_functions("Math")
