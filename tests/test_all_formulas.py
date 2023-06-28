import pytest
import pytest_check as check

from numbers_parser import Document
from numbers_parser.cell import ErrorCell, TextCell, BoolCell, DurationCell

DOCUMENT = "tests/data/test-all-formulas.numbers"
DOCUMENT_13_1 = "tests/data/test-all-formulas.numbers"


def compare_table_functions(sheet_name, filename):
    doc = Document(filename)
    sheet = doc.sheets[sheet_name]
    table = sheet.tables["Tests"]
    for i, row in enumerate(table.rows()):
        if i == 0 or not row[3].value:
            # Skip header and invalid test rows
            continue

        if isinstance(row[0], BoolCell):
            # Test value is true/false
            formula_text = str(row[0].value).upper()
        else:
            formula_text = row[0].value
        formula_result = row[1].value
        formula_ref_value = row[2].value
        check.is_not_none(row[1].formula, f"exists@{i}")
        check.equal(row[1].formula, formula_text, f"formula@{i}")
        if isinstance(row[1], ErrorCell):
            pass
        elif isinstance(row[1], TextCell) and formula_ref_value is None:
            pass
        else:
            check.equal(formula_result, formula_ref_value, f"value@{i}")


def test_information_functions():
    compare_table_functions("Information", DOCUMENT)
    compare_table_functions("Information", DOCUMENT_13_1)


def test_text_functions():
    compare_table_functions("Text", DOCUMENT)
    compare_table_functions("Text", DOCUMENT_13_1)


def test_math_functions():
    compare_table_functions("Math", DOCUMENT)
    compare_table_functions("Math", DOCUMENT_13_1)


def test_reference_functions():
    compare_table_functions("Reference", DOCUMENT)
    compare_table_functions("Reference", DOCUMENT_13_1)


def test_date_functions():
    compare_table_functions("Date", DOCUMENT)
    compare_table_functions("Date", DOCUMENT_13_1)


def test_statistical_functions():
    compare_table_functions("Statistical", DOCUMENT)
    compare_table_functions("Statistical", DOCUMENT_13_1)


def test_extra_functions():
    compare_table_functions("Formulas", "tests/data/test-extra-formulas.numbers")


def test_new_functions():
    compare_table_functions("Formulas", "tests/data/test-new-formulas.numbers")
