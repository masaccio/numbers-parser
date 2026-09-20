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


def test_encrypted_save(configurable_save_file):
    doc = Document("tests/data/encrypted.numbers", password="s3cr3t")  # noqa: S106
    table = doc.sheets[0].tables[0]
    table.write(0, 0, "Encryption")
    doc.save(configurable_save_file, password="r3@llys3cr3t")  # noqa: S106

    new_doc = Document(configurable_save_file, password="r3@llys3cr3t")  # noqa: S106
    new_table = new_doc.sheets[0].tables[0]
    assert new_table.cell(0, 0).value == "Encryption"
