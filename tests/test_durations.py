import pytest
import pytest_check as check


from datetime import datetime
from numbers_parser import Document
from numbers_parser.cell import EmptyCell


def test_duration_formatting():
    doc = Document("tests/data/duration_112.numbers")
    for sheet in doc.sheets:
        table = sheet.tables[0]
        for row in table.iter_rows(min_row=1):
            if type(row[13]) != EmptyCell:
                duration = row[6].formatted_value
                ref = row[13].value
                check.equal(duration, ref)
