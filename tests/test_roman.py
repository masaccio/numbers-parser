import pytest

from numbers_parser.roman import to_roman


def test_to_roman():
    assert to_roman(1) == "I"
    assert to_roman(4) == "IV"
    assert to_roman(9) == "IX"
    assert to_roman(40) == "XL"
    assert to_roman(90) == "XC"
    assert to_roman(400) == "CD"
    assert to_roman(500) == "D"
    assert to_roman(900) == "CM"
    assert to_roman(1000) == "M"
    assert to_roman(3999) == "MMMCMXCIX"
    assert to_roman(0) == "N"

    with pytest.raises(ValueError, match="Number out of range for Roman numerals") as e:
        _ = to_roman(4000)
    assert "Number out of range for Roman numerals" in str(e)
    with pytest.raises(ValueError) as e:
        _ = to_roman(-1)
    assert "Number out of range for Roman numerals" in str(e)
