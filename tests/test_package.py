import pytest

from numbers_parser import Document, FileFormatError


def test_invalid_packages():
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numberz")
    assert "invalid Numbers document (not a .numbers directory)" in str(e)

    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numbers")
    assert "invalid Numbers document (missing files)" in str(e)

    with pytest.warns(RuntimeWarning) as record:
        _ = Document("tests/data/invalid-ver.numbers")
    # assert len(record) == 1
    assert "unsupported version 99.9" in str(record[0])

    doc = Document()
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numberz")
    assert "invalid Numbers document (not a .numbers directory)" in str(e)

    with pytest.raises(FileFormatError) as e:
        doc.save("tests/data/invalid.numbers", package=True)
    assert "folder is not a numbers package" in str(e)

    with pytest.raises(FileFormatError) as e:
        doc.save("tests/data/invalid-props.numbers", package=True)
    assert "invalid Numbers document (missing files)" in str(e)


def test_package_save(configurable_save_file):
    doc = Document()
    table = doc.sheets[0].tables[0]
    table.write(0, 0, "Dog")
    table.write(0, 1, "Cat")
    table.write(1, 0, "Rat")
    table.write(1, 1, "Snake")
    doc.save(configurable_save_file, package=True)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert table.cell(0, 0).value == "Dog"
    assert table.cell(0, 1).value == "Cat"
    assert table.cell(1, 0).value == "Rat"
    assert table.cell(1, 1).value == "Snake"
