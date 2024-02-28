import pytest

from numbers_parser import RGB, Document, FileError, FileFormatError


def test_invalid_packages(configurable_save_file):
    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numberz")
    assert "invalid Numbers document (not a .numbers package/file)" in str(e)

    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/invalid.numbers")
    assert "invalid Numbers document (missing files)" in str(e)

    with pytest.warns(RuntimeWarning) as record:
        _ = Document("tests/data/invalid-ver.numbers")
    # assert len(record) == 1
    assert "unsupported version '99.9'" in str(record[0])

    doc = Document()
    with pytest.raises(FileFormatError) as e:
        _ = doc.save("tests/data/invalid.numberz", package=True)
    assert "invalid Numbers document (not a Numbers package)" in str(e)

    with pytest.raises(FileFormatError) as e:
        doc.save("tests/data/invalid.numbers", package=True)
    assert "folder is not a numbers package" in str(e)

    with pytest.raises(FileFormatError) as e:
        doc.save("tests/data/invalid-props.numbers", package=True)
    assert "invalid Numbers document (missing files)" in str(e)

    with pytest.raises(FileFormatError) as e:
        doc.save(configurable_save_file)
        doc.save(configurable_save_file, package=True)
    assert "cannot overwrite Numbers document file with package" in str(e)

    with pytest.raises(FileFormatError) as e:
        _ = Document("tests/data/corrupted-zip.numbers")
    assert "invalid Numbers document" in str(e)

    with pytest.raises(FileError) as e:
        _ = Document("tests/data/NOT-FOUND.numbers")
    assert "no such file or directory" in str(e)


def test_package_save(configurable_save_file):
    doc = Document("tests/data/test-package.numbers")
    doc.save(configurable_save_file, package=True)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert table.cell(0, 0).style.bg_image.filename == "cat.jpg"
    assert table.cell(0, 1).value == "Cat"
    style = table.cell(1, 1).style
    assert style.font_color == RGB(255, 255, 255)
    assert style.bg_color == RGB(238, 34, 12)
    assert style.bold
