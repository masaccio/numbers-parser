import struct

import pytest
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

from numbers_parser import RGB, Document, FileError
from numbers_parser.iwork import IWork, IWPasswordVerifier


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


def test_password_verifier_rejects_invalid_formats():
    with pytest.raises(ValueError, match="unrecognized format length"):
        IWPasswordVerifier(b"too short")

    verifier_data = struct.pack("<HH I 16s 16s 64s", 1, 1, 1, b"salt" * 4, b"iv" * 8, b"data" * 16)
    with pytest.raises(ValueError, match="Unsupported version or format"):
        IWPasswordVerifier(verifier_data)


def test_password_verifier_password_results():
    work = IWork()
    verifier_data, key = work._generate_verifier_and_key("s3cr3t")
    verifier = IWPasswordVerifier(verifier_data)

    assert verifier.create_key_with_password("") is None
    assert verifier.create_key_with_password("wrong") is None
    assert verifier.create_key_with_password("s3cr3t") == key

    tampered_data = bytearray(verifier_data)
    tampered_data[-1] ^= 1
    assert IWPasswordVerifier(bytes(tampered_data)).create_key_with_password("s3cr3t") is None


def test_iwa_encryption_round_trip():
    work = IWork()
    _, key = work._generate_verifier_and_key("s3cr3t")
    data = b"encrypted IWA payload"

    encrypted = work._encrypt_using_iwa_key(data, key)

    assert work._decrypt_using_iwa_key(encrypted, key) == data


def test_iwa_decryption_rejects_short_data():
    with pytest.raises(ValueError, match="Data too short"):
        IWork()._decrypt_using_iwa_key(b"x" * 35, b"k" * 16)


def test_iwa_decryption_rejects_invalid_padding():
    work = IWork()
    _, key = work._generate_verifier_and_key("s3cr3t")
    encrypted = bytearray(work._encrypt_using_iwa_key(b"payload", key))
    encrypted[-21] ^= 1

    with pytest.raises(ValueError, match="PKCS7 unpadding failed"):
        work._decrypt_using_iwa_key(bytes(encrypted), key)


def test_iwa_decryption_rejects_short_decrypted_payload():
    key = b"k" * 16
    iv = b"i" * 16
    padder = padding.PKCS7(128).padder()
    padded_data = padder.update(b"short") + padder.finalize()
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
    encryptor = cipher.encryptor()
    encrypted_payload = encryptor.update(padded_data) + encryptor.finalize()

    with pytest.raises(ValueError, match="too short to discard"):
        IWork()._decrypt_using_iwa_key(iv + encrypted_payload + b"g" * 20, key)
