import re

# Cell reference conversion from  https://github.com/jmcnamara/XlsxWriter
# Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def xl_cell_to_rowcol(cell_str: str) -> tuple:
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

    Parameters
    ----------
    cell_str:  str
        A1 notation cell reference

    Returns
    -------
    row, col: int, int
        Cell row and column numbers (zero indexed).

    """
    if not cell_str:
        return 0, 0

    match = range_parts.match(cell_str)
    if not match:
        msg = f"invalid cell reference {cell_str}"
        raise IndexError(msg)

    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    col = 0
    for expn, char in enumerate(reversed(col_str)):
        col += (ord(char) - ord("A") + 1) * (26**expn)

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col


def xl_range(first_row, first_col, last_row, last_col):
    """
    Convert zero indexed row and col cell references to a A1:B1 range string.

    Parameters
    ----------
    first_row: int
        The first cell row.
    first_col: int
        The first cell column.
    last_row: int
        The last cell row.
    last_col: int
        The last cell column.

    Returns
    -------
    str:
        A1:B1 style range string.

    """
    range1 = xl_rowcol_to_cell(first_row, first_col)
    range2 = xl_rowcol_to_cell(last_row, last_col)

    if range1 == range2:
        return range1
    return range1 + ":" + range2


def xl_rowcol_to_cell(row, col, row_abs=False, col_abs=False):
    """
    Convert a zero indexed row and column cell reference to a A1 style string.

    Parameters
    ----------
    row: int
         The cell row.
    col: int
        The cell column.
    row_abs: bool
        If ``True``, make the row absolute.
    col_abs: bool
        If ``True``, make the column absolute.

    Returns
    -------
    str:
        A1 style string.

    """
    if row < 0:
        msg = f"row reference {row} below zero"
        raise IndexError(msg)

    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    row += 1  # Change to 1-index.
    row_abs = "$" if row_abs else ""

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)


def xl_col_to_name(col, col_abs=False):
    """
    Convert a zero indexed column cell reference to a string.

    Parameters
    ----------
    col: int
        The column number (zero indexed).
    col_abs: bool, default: False
        If ``True``, make the column absolute.

    Returns
    -------
        str:
            Column in A1 notation.

    """
    if col < 0:
        msg = f"column reference {col} below zero"
        raise IndexError(msg)

    col += 1  # Change to 1-index.
    col_str = ""
    col_abs = "$" if col_abs else ""

    while col:
        # Set remainder from 1 .. 26
        remainder = col % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col = int((col - 1) / 26)

    return col_abs + col_str
