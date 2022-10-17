import pytest

from numbers_parser.numbers_uuid import NumbersUUID


def test_uuid():
    uuid = NumbersUUID()
    assert str(uuid).count("-") == 4

    uuid = NumbersUUID(0xFF00FF00EE00EE00DD00DD00CC00BB00)
    assert uuid.hex == "ff00ff00ee00ee00dd00dd00cc00bb00"
    uuid = NumbersUUID("12345678000000001234567811111111")
    assert uuid.int == 0x12345678000000001234567811111111

    assert NumbersUUID(uuid.protobuf4).int == uuid.int
    assert NumbersUUID(uuid.protobuf4).protobuf2.lower == 0x1234567811111111
    assert NumbersUUID(uuid.protobuf2).int == uuid.int
    assert NumbersUUID(uuid.protobuf2).protobuf4.uuid_w0 == 0x11111111

    uuid = NumbersUUID(
        {"uuid_w0": 0x1234, "uuid_w1": 0xFFFF, "uuid_w2": 0, "uuid_w3": 0x1111}
    )
    assert uuid.hex == "00001111000000000000ffff00001234"

    uuid = NumbersUUID({"upper": 0x1234, "lower": 0xFFFF})
    assert uuid.hex == "0000000000001234000000000000ffff"
