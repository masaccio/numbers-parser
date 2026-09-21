import struct

import pytest
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

from numbers_parser import RGB, Document, FileError
from numbers_parser.iwork import IWorkCrypto


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


def test_encrypted_non_latin_password():
    doc = Document("tests/data/encrypted-non-latin.numbers", password="秘密")  # noqa: S106

    assert doc.sheets[0].tables[0].cell(0, 0).value == "Decryption"


def test_encrypted_save(configurable_save_file):
    doc = Document("tests/data/encrypted.numbers", password="s3cr3t")  # noqa: S106
    table = doc.sheets[0].tables[0]
    table.write(0, 0, "Encryption")
    doc.save(configurable_save_file, password="r3@llys3cr3t")  # noqa: S106

    new_doc = Document(configurable_save_file, password="r3@llys3cr3t")  # noqa: S106
    new_table = new_doc.sheets[0].tables[0]
    assert new_table.cell(0, 0).value == "Encryption"

    with pytest.raises(FileError, match=r"Invalid password\. Hint is 'No hint'"):
        Document(configurable_save_file, password="invalid")  # noqa: S106


def test_encrypted_save_with_hint(configurable_save_file):
    doc = Document("tests/data/encrypted.numbers", password="s3cr3t")  # noqa: S106
    doc.save(configurable_save_file, password="r3@llys3cr3t", hint="Remember me")  # noqa: S106

    with pytest.raises(FileError, match=r"Invalid password\. Hint is 'Remember me'"):
        Document(configurable_save_file, password="invalid")  # noqa: S106


def test_encrypted_package_save(configurable_save_file):
    doc = Document("tests/data/encrypted.numbers", password="s3cr3t")  # noqa: S106
    doc.save(configurable_save_file, package=True, password="r3@llys3cr3t")  # noqa: S106

    assert (configurable_save_file / ".iwph").is_file()
    assert (configurable_save_file / ".iwpv2").is_file()
    new_doc = Document(configurable_save_file, password="r3@llys3cr3t")  # noqa: S106
    assert new_doc.sheets[0].tables[0].cell(0, 0).value == "Decryption"


def test_password_verifier_rejects_invalid_formats():
    with pytest.raises(ValueError, match="unrecognized verifier format length"):
        IWorkCrypto.from_password_verifier(b"too short", "s3cr3t")

    verifier_data = struct.pack("<HH I 16s 16s 64s", 1, 1, 1, b"salt" * 4, b"iv" * 8, b"data" * 16)
    with pytest.raises(ValueError, match="Unsupported version or format"):
        IWorkCrypto.from_password_verifier(verifier_data, "s3cr3t")


def test_password_verifier_password_results():
    verifier_data, _ = IWorkCrypto.from_password("s3cr3t")

    assert IWorkCrypto.from_password_verifier(verifier_data, "") is None
    assert IWorkCrypto.from_password_verifier(verifier_data, "wrong") is None
    assert IWorkCrypto.from_password_verifier(verifier_data, "s3cr3t") is not None

    tampered_data = bytearray(verifier_data)
    tampered_data[-1] ^= 1
    assert IWorkCrypto.from_password_verifier(bytes(tampered_data), "s3cr3t") is None


def test_iwa_encryption_round_trip():
    _, crypto = IWorkCrypto.from_password("s3cr3t")
    data = b"encrypted IWA payload"

    encrypted = crypto.encrypt_iwa(data)

    assert crypto.decrypt_iwa(encrypted) == data


def test_iwa_decryption_rejects_short_data():
    with pytest.raises(ValueError, match="Data too short"):
        IWorkCrypto(b"k" * 16).decrypt_iwa(b"x" * 35)

    with pytest.raises(ValueError, match="Data too short"):
        IWorkCrypto(b"k" * 16).decrypt_iwa(b"x" * 53)


def test_iwa_decryption_rejects_invalid_padding():
    _, crypto = IWorkCrypto.from_password("s3cr3t")
    encrypted = bytearray(crypto.encrypt_iwa(b"payload"))
    encrypted[-21] ^= 1

    with pytest.raises(ValueError, match="PKCS7 unpadding failed"):
        crypto.decrypt_iwa(bytes(encrypted))


def test_iwa_decryption_rejects_short_decrypted_payload():
    crypto = IWorkCrypto(b"k" * 16)
    iv = b"i" * 16
    padder = padding.PKCS7(128).padder()
    padded_data = padder.update(b"short") + padder.finalize()
    cipher = Cipher(algorithms.AES(b"k" * 16), modes.CBC(iv), backend=default_backend())
    encryptor = cipher.encryptor()
    encrypted_payload = encryptor.update(padded_data) + encryptor.finalize()

    with pytest.raises(ValueError, match="too short to discard"):
        crypto.decrypt_iwa(iv + encrypted_payload + b"g" * 20)
