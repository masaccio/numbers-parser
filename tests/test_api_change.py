import pytest

from numbers_parser import Document


def test_api_change():
    doc = Document("tests/data/issue-43.numbers")
    table = doc.sheets[0].tables[0]

    with pytest.warns(DeprecationWarning) as record:
        value = table.cell("A1").image_data
    assert len(record) == 1
    assert len(value) == 87857

    with pytest.warns(DeprecationWarning) as record:
        value = table.cell("A1").image_filename
    assert len(record) == 1
    assert value == "pasted-image-17.png"

    with pytest.warns(DeprecationWarning) as record:
        value = table.cell("C1").image_data
    assert len(record) == 1
    assert value is None

    with pytest.warns(DeprecationWarning) as record:
        value = table.cell("C1").image_filename
    assert len(record) == 1
    assert value is None
