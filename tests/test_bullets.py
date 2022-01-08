import pytest

from numbers_parser import Document, ParagraphTextCell

TEST_REF_1 = "this\nis some\nbulleted\ntext"
TEST_REF_2 = "this\nis some\nnumbered\ntext"
TEST_REF_3 = "this\nis some\nlettered\ntext"


def test_bullets():
    doc = Document("tests/data/test-bullets.numbers")
    sheets = doc.sheets()
    tables = sheets[0].tables()
    table = tables[0]
    assert table.cell(0, 1).value == TEST_REF_1
    assert table.cell(1, 1).value == TEST_REF_2
    assert table.cell(2, 1).value == TEST_REF_3
    assert table.cell(0, 1).paragraphs[0] == TEST_REF_1.split("\n")[0] + "\n"
    assert table.cell(0, 1).paragraphs[2] == TEST_REF_1.split("\n")[2] + "\n"
    assert table.cell(0, 1).paragraphs[3] == TEST_REF_1.split("\n")[3]
