from random import randint

import pytest

from numbers_parser import RGB, Alignment, Border, Document, xl_rowcol_to_cell


def test_many_rows():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet_1"].tables
    data = tables[0].rows(values_only=True)
    ref = []
    _ = [ref.append(["ROW" + str(x)]) for x in range(1, 601)]
    assert data == ref


def test_many_columns():
    doc = Document("tests/data/test-3.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet_2"].tables
    data = tables[0].rows(values_only=True)
    ref = [["COLUMN" + str(x) for x in range(1, 601)]]
    assert data == ref


def ref_cell_text(row, col):
    return "CELL [" + str(row + 1) + "," + str(col + 1) + "]"


def test_large_table():
    row = [None for i in range(270)]
    ref = [row.copy() for i in range(90)]
    for i in range(90):
        ref[i][i] = ref_cell_text(i, i)
        ref[i][90 + i] = ref_cell_text(i, 90 + i)
        ref[i][180 + i] = ref_cell_text(i, 180 + i)

    doc = Document("tests/data/test-6.numbers")
    sheets = doc.sheets
    tables = sheets["Sheet"].tables
    data = tables[0].rows(values_only=True)
    assert data == ref


@pytest.mark.experimental
def test_profiling_tables(configurable_save_file):
    doc = Document()
    for _ in range(10):
        doc.add_sheet()
    for sheet in doc.sheets:
        for _ in range(10):
            sheet.add_table()
    doc.save(configurable_save_file)


@pytest.mark.experimental
def test_profiling_borders(configurable_save_file):
    doc = Document()
    table = doc.default_table
    for row in range(100):
        for col in range(100):
            color = RGB(randint(0, 256), randint(0, 256), randint(0, 256))  # noqa: S311
            border = Border(color=color, width=2.0, style="solid")
            table.write(row, col, xl_rowcol_to_cell(row, col))
            table.set_cell_border(row, col, ["left", "right", "top", "bottom"], border)
    doc.save(configurable_save_file)


@pytest.mark.experimental
def test_profiling_styles(configurable_save_file):
    doc = Document()
    style_idx = 1
    table = doc.default_table
    for row in range(100):
        for col in range(100):
            color = RGB(randint(0, 256), randint(0, 256), randint(0, 256))  # noqa: S311
            style = doc.add_style(
                name=f"Red Text {style_idx}",
                font_color=color,
                bold=True,
                italic=True,
                alignment=Alignment("center", "middle"),
            )
            style_idx += 1
            table.write(row, col, xl_rowcol_to_cell(row, col), style=style)

    doc.save(configurable_save_file)
