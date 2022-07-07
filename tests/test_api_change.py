import pytest

from numbers_parser import Document


def test_api_change():
    doc = Document("tests/data/test-1.numbers")

    with pytest.warns(DeprecationWarning) as record:
        sheets = doc.sheets()
    assert len(record) == 1
    assert "sheets() is deprecated and will be removed" in record[0].message.args[0]
    assert sheets[0].name == "ZZZ_Sheet_1"

    sheet = doc.sheets[0]
    with pytest.warns(DeprecationWarning) as record:
        tables = sheet.tables()
    assert len(record) == 1
    assert "tables() is deprecated and will be removed" in record[0].message.args[0]
    assert tables[0].name == "ZZZ_Table_1"

    assert len(doc.sheets) == 2
    assert len(doc.sheets[0].tables) == 2
    assert doc.sheets[0].tables[0].cell(1, 0).value == "YYY_ROW_1"
