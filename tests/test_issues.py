import pytest

from numbers_parser import Document, NumberCell
from datetime import datetime, timedelta

ISSUE_3_REF = [("A", "B"), (2.0, 0.0), (3.0, 1.0), (None, None)]
ISSUE_4_REF_1 = "Part 1 \n\nPart 2\n"
ISSUE_4_REF_2 = "\n\nPart 1 \n\n\nPart 2\n\n"
ISSUE_4_REF_3 = "今天是个好日子"
ISSUE_4_REF_4 = "Lorem ipsum\n\ndolor sit amet,\n\nconsectetur adipiscing"

ISSUE_7_REF_1 = """Open http://www.mytest.com/music on Desktop. Click Radio on left penal
Take a screenshot including the bottom banner"""
ISSUE_7_REF_2 = """Click the bottom banner
See the generic Individual upsell
Take a screenshot
Dismiss the upsell"""

ISSUE_10_REF = [
    123,
    12.34,  # US$
    12.34,  # percentage
    12.34,  # 12 17/50
    1234,  # 0x42d
    1234.56,
    "123",
    datetime(2021, 4, 3, 0, 0, 0),
    timedelta(days=4, hours=2, minutes=3),
    timedelta(days=5, hours=4, minutes=3, seconds=20),
    timedelta(hours=4, minutes=3, seconds=2, milliseconds=10),
    timedelta(weeks=12, hours=5),
    True,  # Checkbox
    3,
    50,
    12,
    "Item 1",
    123.4567,  # Formatted as 123.46
]


def test_issue_3():
    doc = Document("tests/data/issue-3.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    ref = []
    for row in tables[0].iter_rows():
        ref.append(tuple([x.value for x in row]))
    assert ref == ISSUE_3_REF

    ref = []
    for row in tables[0].iter_rows(values_only=True):
        ref.append(row)
    assert ref == ISSUE_3_REF


def test_issue_4():
    doc = Document("tests/data/issue-4.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    assert table.cell(1, 0).value == ISSUE_4_REF_1
    assert table.cell(1, 1).value == ISSUE_4_REF_2
    assert table.cell(2, 0).value == ISSUE_4_REF_3
    assert table.cell(2, 1).value == ISSUE_4_REF_4


def test_issue_7():
    doc = Document("tests/data/issue-7.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    assert table.cell(1, 1).value == ISSUE_7_REF_1
    assert table.cell(2, 1).value == ISSUE_7_REF_2
    table.cell(1, 1).bullets[0] == ISSUE_7_REF_1.split("\n")[0] + "\n"
    table.cell(2, 1).bullets[2] == ISSUE_7_REF_2.split("\n")[2] + "\n"


def test_issue_9():
    doc = Document("tests/data/issue-9.numbers")
    sheets = doc.sheets()
    assert len(sheets) == 7


def test_issue_10():
    doc = Document("tests/data/issue-10.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    for i, test_value in enumerate(ISSUE_10_REF):
        assert table.cell(i + 1, 1).value == test_value


def test_issue_14():
    doc = Document("tests/data/issue-14.numbers")
    sheets = doc.sheets()
    table = sheets["Ex 1"].tables()[0]
    assert table.cell("G2").value == 19
    assert table.cell("G2").formula == "XLOOKUP(F2,A2:A15,B2:B15)"
    table = sheets["Ex 6"].tables()[0]
    assert table.cell("F2").value == "Pam"
    assert table.cell("F2").formula == "XLOOKUP(F1,$B$2:$B$15,$A$2:$A$15,,,-1)"
