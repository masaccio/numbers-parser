import pytest

from numbers_parser import Document


def test_read_folder():
    doc = Document("tests/data/test-1.numbers")

    with pytest.warns(DeprecationWarning) as record:
        _ = doc.sheets()
    assert len(record) == 1
    assert "sheets() is deprecated and will be removed" in record[0].message.args[0]

    sheet = doc.sheets[0]
    with pytest.warns(DeprecationWarning) as record:
        _ = sheet.tables()
    assert len(record) == 1
    assert "tables() is deprecated and will be removed" in record[0].message.args[0]

    assert len(doc.sheets) == 2
    assert len(doc.sheets[0].tables) == 2
    assert doc.sheets[0].tables[0].cell(1, 0).value == "YYY_ROW_1"
