def to_roman(value: int) -> str:
    """Convert an integer to a Roman number in the available range of values."""
    if value == 0:
        return "N"

    if value < 1 or value > 3999:
        raise ValueError("Number out of range for Roman numerals")

    int_values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    roman_symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]

    roman_num = ""
    for i in range(len(int_values)):
        while value >= int_values[i]:
            roman_num += roman_symbols[i]
            value -= int_values[i]

    return roman_num
