import pytest

from datetime import datetime
from numbers_parser import Document


STATEMENT_REF = [
    ["Date", "Transaction Details", "Paid In", "Withdrawn", "Balance", "Category" ],
    [datetime(2020, 5, 1), "Debit to Marlon Computing", 250.00, None, 2021.50, "Loans" ],
    [datetime(2020, 5, 1), "Parking Downtown 932891", None, 2.99, 2018.51, "Entertainment" ],
    [datetime(2020, 5, 1), "Grant Coffee", None, 13.2, 2005.31, "Eating Out" ],
]

def test_read_folder():
    doc = Document("tests/data/test-2.numbers")
    sheets = doc.sheets()
    tables = sheets["Account"].tables()
    data = tables["Statement"].data
    assert data == STATEMENT_REF
