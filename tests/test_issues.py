import magic

from numbers_parser import Document, NumberCell
from pendulum import datetime, duration
from numbers_parser.cell import ErrorCell, EmptyCell

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
    table.cell(1, 1).bullets[0] == ISSUE_7_REF_1.split("\n")[0] + "\n"
    table.cell(2, 1).bullets[2] == ISSUE_7_REF_2.split("\n")[2] + "\n"


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
    doc = Document("tests/data/issue-17.numbers")
    sheets = doc.sheets
    table = sheets[0].tables[0]
    assert table.cell(0, 0).value == 123.0
    assert table.cell(0, 0).is_merged == False
    assert table.cell(0, 0).formula == None


def test_issue_18():
    doc = Document("tests/data/issue-18.numbers")
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
        "cat-numbers",
        "--brief",
        "tests/data/issue-42.numbers",
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    assert lines[5] == ",#REF!"
    assert lines[6] == "7.0,#REF!"

    ret = script_runner.run(
        "cat-numbers",
        "--brief",
        "--formulas",
        "tests/data/issue-42.numbers",
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
    assert type(cell) == EmptyCell
    assert "PNG image data" in magic.from_buffer(cell.image_data)
    assert len(cell.image_data) == 87857
    assert cell.image_filename == "pasted-image-17.png"

    cell = table.cell("B1")
    assert cell.value == "text "
    assert "TIFF image data" in magic.from_buffer(cell.image_data)
    assert len(cell.image_data) == 365398
    assert cell.image_filename == "pasted-image-19.tiff"

    assert table.cell("C1").image_data is None
    assert table.cell("C1").image_filename is None


def test_issue_44(script_runner):
    doc = Document("tests/data/issue-44.numbers")
    table = doc.sheets[0].tables[0]
    for row_num, ref in enumerate(ISSUE_44_REF):
        assert table.cell(row_num, 0).formatted_value == ref[0]
        assert table.cell(row_num, 1).formatted_value == ref[1]

    ret = script_runner.run(
        "cat-numbers",
        "--brief",
        "--formatting",
        "tests/data/issue-44.numbers",
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.success
    lines = ret.stdout.split("\r\n")
    for row_num, ref in enumerate(ISSUE_44_REF):
        # Remove " escapes
        ref_str = ",".join([x.replace('"', "") for x in ref])
        test_str = lines[row_num].replace('"', "")
        assert test_str == ref_str


def test_issue_49(tmp_path, pytestconfig):
    doc = Document("tests/data/issue-49.numbers")

    if pytestconfig.getoption("save_file") is not None:
        new_filename = pytestconfig.getoption("save_file")
    else:
        new_filename = tmp_path / "test-1-new.numbers"
    doc.save(new_filename)

    doc = Document(new_filename)
    table = doc.sheets[0].tables[0]
    test_values = [int(table.cell(row_num, 1).value) for row_num in range(1, 6)]
    assert test_values == [100, 50, 0, -50, -100]
