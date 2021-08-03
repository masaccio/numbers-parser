import pytest
import pytest_check as check

from numbers_parser import Document
from numbers_parser.cell import ErrorCell, TextCell, BoolCell

DOCUMENT = "tests/data/test-all-forumulas.numbers"


def compare_table_functions(sheet_name, filename=DOCUMENT):
    doc = Document(filename)
    sheet = doc.sheets()[sheet_name]
    table = sheet.tables()["Tests"]
    for i, row in enumerate(table.rows()):
        if i == 0 or not row[3].value:
            # Skip header and invalid test rows
            continue

        if isinstance(row[0], BoolCell):
            formula_text = str(row[0].value).upper()
        else:
            formula_text = row[0].value
        formula_result = row[1].value
        formula_ref_value = row[2].value
        check.is_not_none(row[1].formula)
        check.equal(row[1].formula, formula_text)
        if isinstance(row[1], ErrorCell):
            pass
        elif isinstance(row[1], TextCell) and formula_ref_value is None:
            pass
        else:
            check.equal(formula_result, formula_ref_value)


def test_information_functions():
    compare_table_functions("Information")


def test_text_functions():
    compare_table_functions("Text")


def test_math_functions():
    compare_table_functions("Math")


def test_reference_functions():
    compare_table_functions("Reference")


@pytest.mark.experimental
def test_date_functions():
    compare_table_functions("Date")


def test_statistical_functions():
    compare_table_functions("Statistical")


def test_extra_functions():
    compare_table_functions("Formulas", "tests/data/test-extra-formulas.numbers")
