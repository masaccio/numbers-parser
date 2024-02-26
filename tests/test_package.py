import pytest

from numbers_parser import Document, FileFormatError


def test_invalid_packages():
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numberz")
    assert "Invalid Numbers file (not a .numbers directory)" in str(e)

    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numbers")
    assert "Invalid Numbers file (missing files)" in str(e)

    with pytest.warns(RuntimeWarning) as record:
        _ = Document("tests/data/invalid-ver.numbers")
    assert len(record) == 1
    assert "unsupported version 99.9" in str(record[0])
