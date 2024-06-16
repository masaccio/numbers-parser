from sys import version_info

import magic
import pytest
from pendulum import datetime, duration

from numbers_parser import (
    RGB,
    BackgroundImage,
    Border,
    Document,
    EmptyCell,
    ErrorCell,
    UnsupportedWarning,
)
from numbers_parser.constants import (
    DEFAULT_COLUMN_COUNT,
    DEFAULT_COLUMN_WIDTH,
    DEFAULT_ROW_COUNT,
    DEFAULT_ROW_HEIGHT,
)

ISSUE_3_REF = [("A", "B"), (2.0, 0.0), (3.0, 1.0), (None, None)]
ISSUE_4_REF_1 = "Part 1 \n\nPart 2\n"
ISSUE_4_REF_2 = "\n\nPart 1 \n\n\nPart 2\n\n"
ISSUE_4_REF_3 = "今天是个好日子"
ISSUE_4_REF_4 = "Lorem ipsum\n\ndolor sit amet,\n\nconsectetur adipiscing"

ISSUE_7_REF_1 = """Open http://www.mytest.com/music on Desktop. Click Radio on left penal
Take a screenshot including the bottom banner"""
ISSUE_7_REF_2 = """Click the bottom banner
See the generic Individual upsell
Take a screenshot
Dismiss the upsell"""

ISSUE_10_REF = [
    123,
    12.34,  # US$
    12.34,  # percentage
    12.34,  # 12 17/50
    1234,  # 0x42d
    1234.56,
    "123",
    datetime(2021, 4, 3, 0, 0, 0),
    duration(days=4, hours=2, minutes=3),
    duration(days=5, hours=4, minutes=3, seconds=20),
    duration(hours=4, minutes=3, seconds=2, milliseconds=10),
    duration(weeks=12, hours=5),
    True,  # Checkbox
    3,
    50,
    12,
    "Item 1",
    123.456789,  # Formatted as 123.46
]

ISSUE_37_REF = [
    ["0:00", "0:00:00"],
    ["1:01", "1:01:01"],
    ["9:09", "9:09:09"],
    ["10:10", "10:10:10"],
    ["11:11", "11:11:11"],
    ["12:12", "12:12:12"],
    ["23:23", "23:23:23"],
]

ISSUE_44_REF = [
    ["1", "1.0000"],
    ["0", "0.0000"],
    ["-2", "-2.00000"],
    ["-100", "100"],
    ["(100)", "(100)"],
    ["-100.1234", "100.1234"],
    ["(100.1234)", "(100.1234)"],
    ["TRUE", "FALSE"],
    ["10000", "10,000"],
    ["1000.125", "1000.12"],
]


def test_issue_3():
    doc = Document("tests/data/issue-3.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    ref = []
    for row in tables[0].iter_rows():
        ref.append(tuple([x.value for x in row]))
    assert ref == ISSUE_3_REF

    ref = []
    for row in tables[0].iter_rows(values_only=True):
        ref.append(row)
    assert ref == ISSUE_3_REF


def test_issue_4():
    doc = Document("tests/data/issue-4.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    assert table.cell(1, 0).value == ISSUE_4_REF_1
    assert table.cell(1, 1).value == ISSUE_4_REF_2
    assert table.cell(2, 0).value == ISSUE_4_REF_3
    assert table.cell(2, 1).value == ISSUE_4_REF_4


def test_issue_7():
    doc = Document("tests/data/issue-7.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    assert table.cell(1, 1).value == ISSUE_7_REF_1
    assert table.cell(2, 1).value == ISSUE_7_REF_2
    assert table.cell(1, 1).bullets[0] == ISSUE_7_REF_1.split("\n")[0]
    assert table.cell(2, 1).bullets[2] == ISSUE_7_REF_2.split("\n")[2]


def test_issue_9():
    doc = Document("tests/data/issue-9.numbers")
    sheets = doc.sheets
    assert len(sheets) == 7


def test_issue_10():
    doc = Document("tests/data/issue-10.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    table = tables[0]
    for i, test_value in enumerate(ISSUE_10_REF):
        assert table.cell(i + 1, 1).value == test_value


def test_issue_14():
    doc = Document("tests/data/issue-14.numbers")
    sheets = doc.sheets
    table = sheets["Ex 1"].tables[0]
    assert table.cell("G2").value == 19
    assert table.cell("G2").formula == "XLOOKUP(F2,A2:A15,B2:B15)"
    table = sheets["Ex 6"].tables[0]
    assert table.cell("F2").value == "Pam"
    assert table.cell("F2").formula == "XLOOKUP(F1,$B$2:$B$15,$A$2:$A$15,,,-1)"


def test_issue_17():
    with pytest.warns(RuntimeWarning) as record:
        doc = Document("tests/data/issue-17.numbers")
    assert len(record) == 2
    assert "can't read Numbers version from document" in str(record[0].message)
    assert "unsupported version ''" in str(record[1].message)
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.cell(0, 0).value == 123.0
    assert not table.cell(0, 0).is_merged
    assert table.cell(0, 0).formula is None


def test_issue_18():
    with pytest.warns(RuntimeWarning) as record:
        doc = Document("tests/data/issue-18.numbers")
    assert len(record) == 2
    assert "can't read Numbers version from document" in str(record[0].message)
    assert "unsupported version ''" in str(record[1].message)
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.merge_ranges == ["B3:D3"]


def test_issue_32():
    doc = Document("tests/data/issue-32.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.cell("A3").value == "Foo"
    assert table.cell("D4").value == 3


def test_issue_35():
    doc = Document("tests/data/issue-35.numbers")
    table = doc.sheets[0].tables[0]
    assert table.cell("A1").value == 72
    assert table.cell("ALL3").value == 62


def test_issue_37():
    doc = Document("tests/data/issue-37.numbers")
    table = doc.sheets[0].tables[0]
    for i, row in enumerate(table.rows()[1:]):
        assert row[-2].formatted_value == ISSUE_37_REF[i][0]
        assert row[-1].formatted_value == ISSUE_37_REF[i][1]


def test_issue_42(script_runner):
    doc = Document("tests/data/issue-42.numbers")
    table = doc.sheets[0].tables[0]
    assert type(table.cell(6, 1)) == ErrorCell
    assert table.cell(3, 1).formula == "#REF!×A4:A6"
    assert table.cell(4, 1).formula == "#REF!×A5:A6"

    ret = script_runner.run(
        ["cat-numbers", "--brief", "tests/data/issue-42.numbers"],
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    assert lines[5] == ",#REF!"
    assert lines[6] == "7.0,#REF!"

    ret = script_runner.run(
        ["cat-numbers", "--brief", "--formulas", "tests/data/issue-42.numbers"],
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    assert lines[4] == "3.0,#REF!×A5:A6"
    assert lines[5] == ",#REF!×A6:A6"
    assert lines[6] == "SUM(A),PRODUCT(B)"


def test_issue_43():
    doc = Document("tests/data/issue-43.numbers")
    table = doc.sheets[0].tables[0]
    cell = table.cell("A1")
    assert isinstance(cell, EmptyCell)
    assert "PNG image data" in magic.from_buffer(cell.style.bg_image.data)
    assert len(cell.style.bg_image.data) == 87857
    assert cell.style.bg_image.filename == "pasted-image.png"

    cell = table.cell("B1")
    assert cell.value == "text "
    assert "TIFF image data" in magic.from_buffer(cell.style.bg_image.data)
    assert len(cell.style.bg_image.data) == 365398
    assert cell.style.bg_image.filename == "pasted-image.tiff"

    assert table.cell("C1").style.bg_image is None


def test_issue_44(script_runner):
    doc = Document("tests/data/issue-44.numbers")
    table = doc.sheets[0].tables[0]
    for row, ref in enumerate(ISSUE_44_REF):
        assert table.cell(row, 0).formatted_value == ref[0]
        assert table.cell(row, 1).formatted_value == ref[1]

    ret = script_runner.run(
        ["cat-numbers", "--brief", "--formatting", "tests/data/issue-44.numbers"],
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    for row, ref in enumerate(ISSUE_44_REF):
        # Remove " escapes
        ref_str = ",".join([x.replace('"', "") for x in ref])
        test_str = lines[row].replace('"', "")
        assert test_str == ref_str


def test_issue_49(configurable_save_file):
    doc = Document("tests/data/issue-49.numbers")
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    test_values = [int(table.cell(row, 1).value) for row in range(1, 6)]
    assert test_values == [100, 50, 0, -50, -100]


def test_issue_50():
    doc = Document("tests/data/issue-50.numbers")
    table = doc.sheets[0].tables[0]
    assert table.num_rows == 65553
    assert table.cell(65552, 0).value == "string 262161"


def test_issue_51(script_runner):
    doc = Document("tests/data/issue-51.numbers")

    table = doc.sheets[0].tables[0]
    for row in table.iter_rows(min_row=1):
        assert str(row[0].formatted_value) == row[1].value

    with pytest.warns(RuntimeWarning) as record:
        table.write(5, 0, 0.00016450000000000001)
    assert len(record) == 1
    assert str(record[0].message) == "'0.00016450000000000001' rounded to 15 significant digits"

    ret = script_runner.run(
        ["cat-numbers", "-b", "tests/data/issue-51.numbers"],
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    assert lines[1:6] == [
        "0.0001645,0.0001645",
        "0.123,0.123",
        "0.1,0.1",
        "-0.0001645,-0.0001645",
        "0.0001645,0.00016450000000000000",
    ]


def test_issue_54():
    doc = Document("tests/data/issue-54.numbers")
    assert doc.sheets[0].tables[0].cell(4, 3).formula == "SUM(Table 1::C1:C4)"
    assert doc.sheets[0].tables[0].cell(4, 2).formula == "SUM(Sheet 2::Table 1::C1:C2)"
    assert (
        doc.sheets[0].tables[1].cell(4, 1).formula
        == "SUM(A1,Sheet 2::Table 1::B1,Sheet 2::Table 1::B3:B4)"
    )
    assert doc.sheets[0].tables[0].cell(4, 2).formula == "SUM(Sheet 2::Table 1::C1:C2)"

    table = doc.sheets[1].tables[0]
    for col in range(9):
        assert table.cell(4, col).formula == table.cell(5, col).value
        assert table.cell(6, col).formula == table.cell(7, col).value


def test_issue_56(tmp_path):
    doc = Document("tests/data/issue-56.numbers")
    table = doc.sheets[0].tables[0]
    assert table.cell("A2").style.bg_color == (255, 149, 202)
    assert table.cell("B2").style.alignment.horizontal.name == "RIGHT"
    assert table.cell("B2").style.alignment.vertical.name == "TOP"

    new_filename = tmp_path / "issue-56-new.numbers"
    doc.save(new_filename)
    print(new_filename)

    new_doc = Document(new_filename)
    new_table = new_doc.sheets[0].tables[0]
    assert new_table.col_width(0) == 98
    assert new_table.col_width(1) == 162
    assert new_table.cell("A2").style.font_name == "Helvetica Neue"
    assert new_table.cell("A2").style.bg_color == (255, 149, 202)
    assert new_table.cell("B2").style.alignment.horizontal.name == "RIGHT"
    assert new_table.cell("B2").style.alignment.vertical.name == "TOP"


def test_issue_59():
    from numbers_parser import Document

    doc = Document("tests/data/issue-59.numbers")
    sheets = doc.sheets
    tables = sheets[0].tables
    assert tables[0].cell("J2").value == "Saturday, 15 May 2021"


@pytest.mark.parametrize("configurable_multi_save_file", [[2]], indirect=True)
def test_issue_60(configurable_multi_save_file):
    test_1_filename, test_2_filename = configurable_multi_save_file

    doc = Document("tests/data/issue-60.numbers")
    doc.save(test_1_filename)

    table = doc.sheets[0].tables[0]
    test_1_doc = Document(test_1_filename)
    test_table_1 = test_1_doc.sheets[0].tables[0]

    doc = Document("tests/data/issue-60.numbers")
    table = doc.sheets[0].tables[0]
    bg_color = table.cell("A2").style.bg_color
    table.write("A2", "Item 2")
    table.cell("A2").style.bg_color = bg_color
    doc.save(test_2_filename)
    test_2_doc = Document(test_2_filename)
    test_table_2 = test_2_doc.sheets[0].tables[0]

    for row, cells in enumerate(table.iter_rows()):
        for col, cell in enumerate(cells):
            for attr in [
                "alignment",
                "bg_image",
                "bg_color",
                "font_color",
                "font_size",
                "font_name",
                "bold",
                "italic",
                "strikethrough",
                "underline",
                "name",
            ]:
                ref = getattr(cell.style, attr)
                assert getattr(test_table_1.cell(row, col).style, attr) == ref
                assert getattr(test_table_2.cell(row, col).style, attr) == ref


def test_issue_66():
    doc = Document("tests/data/issue-66-collab.numbers")
    sheets = doc.sheets
    assert sheets[0].name == "Credit"


def test_issue_69():
    doc = Document("tests/data/issue-69b.numbers")
    table = doc.sheets[0].tables[0]
    assert table.cell(0, 0).style.bg_image.filename == "numbers_1.png"
    assert table.cell(1, 0).style.bg_image.filename == "numbers_2.png"
    assert table.cell(2, 0).style.bg_image.filename == "numbers_3.png"
    assert len(table.cell(0, 0).style.bg_image.data) == 10608
    assert len(table.cell(1, 0).style.bg_image.data) == 19269
    assert len(table.cell(2, 0).style.bg_image.data) == 19256

    doc = Document("tests/data/issue-69.numbers")
    table = doc.sheets[0].tables[0]
    if version_info.minor >= 11:
        assert table.cell(0, 0).style.bg_image.filename == "sssssss的副本.jpeg"
        assert table.cell(9, 4).style.bg_image.filename == "sssssss的副本.jpeg"
    else:
        with pytest.warns(RuntimeWarning) as record:
            assert table.cell(0, 0).style.bg_image is None
        assert len(record) == 1
        assert str(record[0].message) == "Cannot find file 'sssssss的副本.jpeg' in Numbers archive"


def test_issue_73(configurable_save_file):
    doc = Document("tests/data/issue-73.numbers")
    with pytest.warns(UnsupportedWarning) as record:
        doc.save(configurable_save_file)
    assert str(record[0].message) == "Not modifying pivot table 'Table 1 Pivot'"


def test_issue_74_1(configurable_save_file):
    doc = Document("tests/data/test-issue-74.numbers")
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    assert doc.sheets[0].tables[0].cell(0, 1)._control_id is not None


def test_issue_74_2():
    doc = Document()
    table = doc.sheets[0].tables[0]

    table.write(1, 0, "Dog")
    table.set_cell_formatting(1, 0, "popup", popup_values=["Cat", "Dog", "Rabbit"], allow_none=True)


def test_issue_75(configurable_save_file):
    doc = Document("tests/data/test-issue-75.numbers")
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    table = doc.sheets[0].tables[0]
    assert all([table.cell(1, col).style.bg_image.filename == "cat.jpg" for col in range(1, 6)])
    assert all([table.cell(2, col).style.bg_image.filename == "cat.jpg" for col in range(1, 6)])


def test_issue_76(configurable_save_file):
    doc = Document("tests/data/test-issue-76.numbers")
    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    assert "Metadata/Properties.plist" in doc._model.objects._file_store
    assert doc._model.objects._file_store["Metadata/Properties.plist"][0:6] == b"bplist"


def test_issue_77():
    doc = Document("tests/data/issue-77.numbers")
    sheet = doc.sheets[0]
    assert sheet.tables[0].merge_ranges == ["C1:D2"]
    assert sheet.tables[1].merge_ranges == ["C1:D2"]


def test_issue_78(configurable_save_file):
    doc = Document()
    doc.sheets[0].add_table()
    table1 = doc.sheets[0].tables[0]
    table2 = doc.sheets[0].tables[1]

    image_filename = "tests/data/cat.jpg"
    image_data = open(image_filename, mode="rb").read()
    bg_image = BackgroundImage(image_data, image_filename)
    cats_bg_style = doc.add_style(bg_image=bg_image)

    table1.write(0, 0, "cats", style=cats_bg_style)
    table2.write(0, 0, "cats", style=cats_bg_style)

    doc.save(configurable_save_file)

    doc = Document(configurable_save_file)
    assert (
        doc.sheets[0].tables[0].cell(0, 0).style.bg_image.data
        == doc.sheets[0].tables[1].cell(0, 0).style.bg_image.data
    )

    # Coverage for Github debug log guard
    assert doc._model.sheet_name(99999999) is None


@pytest.mark.experimental
def test_issue_80():
    doc = Document("tests/data/issue-80.numbers")
    assert doc.default_table.cell("B8").value == "AB"


def test_issue_83():
    doc = Document()
    table = doc.default_table
    table.write(0, 0, "test")
    with pytest.raises(TypeError) as e:
        table.set_cell_formatting(0, 0, "text")
    assert "unsuported cell format type 'text'" in str(e)


def test_issue_85():
    doc = Document()
    table = doc.default_table

    assert table.num_rows == DEFAULT_ROW_COUNT
    assert table.num_cols == DEFAULT_COLUMN_COUNT
    assert table.height == DEFAULT_ROW_COUNT * DEFAULT_ROW_HEIGHT
    assert table.width == DEFAULT_COLUMN_COUNT * DEFAULT_COLUMN_WIDTH

    table.add_row()
    table.add_column()

    assert table.num_rows == DEFAULT_ROW_COUNT + 1
    assert table.num_cols == DEFAULT_COLUMN_COUNT + 1
    assert table.height == (DEFAULT_ROW_COUNT + 1) * DEFAULT_ROW_HEIGHT
    assert table.width == (DEFAULT_COLUMN_COUNT + 1) * DEFAULT_COLUMN_WIDTH

    table.write(25, 25, "TEST")

    assert table.num_rows == 26
    assert table.num_cols == 26
    assert table.height == 26 * DEFAULT_ROW_HEIGHT
    assert table.width == 26 * DEFAULT_COLUMN_WIDTH

    doc = Document(num_cols=4, num_rows=4)
    sheet = doc.sheets[0]
    table0 = sheet.tables[0]

    assert table0.row_height(0) == 20.0
    assert table0.row_height(2) == 20.0
    assert table0.col_width(0) == 98.0
    assert table0.col_width(1) == 98.0

    for row in range(0, table0.num_rows, 2):
        for col in range(0, table0.num_cols, 2):
            border_style = Border(1 + (2.0 * col), RGB(29, 177, 0), "solid")
            table0.set_cell_border(row, col, "top", border_style)
            table0.set_cell_border(row, col, "left", border_style)
            table0.set_cell_border(row, col, "bottom", border_style)
            table0.set_cell_border(row, col, "right", border_style)

    assert table0.row_height(0) == 25.0
    assert table0.row_height(3) == 22.0
    assert table0.col_width(0) == 99.0
    assert table0.col_width(1) == 101
    assert table0.col_width(2) == 103
    assert table0.col_width(3) == 100
