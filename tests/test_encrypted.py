import pytest

from numbers_parser import RGB, Document, FileError


def test_encrypted_read():
    doc = Document("tests/data/encrypted.numbers", password="s3cr3t")  # noqa: S106
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.cell(0, 0).value == "Decryption"
    assert table.cell(0, 0).style.bold
    assert table.cell(0, 1).style.italic
    assert table.cell(1, 1).style.font_color == RGB(255, 0, 0)

    with pytest.raises(FileError) as e:
        _ = Document("tests/data/encrypted.numbers", password="invalid")  # noqa: S106
    assert "Invalid password. Hint is 's3cr3t'" in str(e)
