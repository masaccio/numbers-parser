def to_roman(value: int) -> str:
    """Convert an integer to a Roman number in the available range of values."""
    if value == 0:
        return "N"

    if value < 1 or value > 3999:
        raise ValueError("Number out of range for Roman numerals")

    roman_map = {
        1000: "M",
        900: "CM",
        500: "D",
        400: "CD",
        100: "C",
        90: "XC",
        50: "L",
        40: "XL",
        10: "X",
        9: "IX",
        5: "V",
        4: "IV",
        1: "I",
    }

    roman_num = ""
    for int_value, roman_symbol in roman_map.items():
        while value >= int_value:
            roman_num += roman_symbol
            value -= int_value

    return roman_num
