import pytest
import roman

from numbers_parser.roman import to_roman


def test_roman():
    for num in range(1, 4000):
        assert to_roman(num) == roman.toRoman(num)

    with pytest.raises(ValueError) as e:
        _ = to_roman(4000)
    assert "Number out of range for Roman numerals" in str(e)
    with pytest.raises(ValueError) as e:
        _ = to_roman(-1)
    assert "Number out of range for Roman numerals" in str(e)

    assert to_roman(0) == "N"
