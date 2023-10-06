import csv

import pytest
import pytest_check as check

from numbers_parser import __version__
from numbers_parser.experimental import _enable_experimental_features

ZZZ_TABLE_1_REF = [
    [None, "YYY_COL_1", "YYY_COL_2"],
    ["YYY_ROW_1", "YYY_1_1", "YYY_1_2"],
    ["YYY_ROW_2", "YYY_2_1", "YYY_2_2"],
    ["YYY_ROW_3", "YYY_3_1", "YYY_3_2"],
    ["YYY_ROW_4", "YYY_4_1", "YYY_4_2"],
]

ZZZ_TABLE_2_REF = [
    [None, "ZZZ_COL_1", "ZZZ_COL_2", "ZZZ_COL_3"],
    ["ZZZ_ROW_1", "ZZZ_1_1", "ZZZ_1_2", "ZZZ_1_3"],
    ["ZZZ_ROW_2", "ZZZ_2_1", "ZZZ_2_2", "ZZZ_2_3"],
    ["ZZZ_ROW_3", "ZZZ_3_1", "ZZZ_3_2", "ZZZ_3_3"],
]

XXX_TABLE_1_REF = [
    [None, "XXX_COL_1", "XXX_COL_2", "XXX_COL_3", "XXX_COL_4", "XXX_COL_5"],
    ["XXX_ROW_1", "XXX_1_1", "XXX_1_2", "XXX_1_3", "XXX_1_4", "XXX_1_5"],
    ["XXX_ROW_2", "XXX_2_1", "XXX_2_2", None, "XXX_2_4", "XXX_2_5"],
    ["XXX_ROW_3", "XXX_3_1", None, "XXX_3_3", "XXX_3_4", "XXX_3_5"],
]

DOCUMENT = "tests/data/test-1.numbers"


@pytest.mark.script_launch_mode("subprocess")
def test_no_documents(script_runner):
    ret = script_runner.run(["cat-numbers", "--brief"], print_result=False)
    assert ret.success
    assert "usage: cat-numbers" in ret.stdout
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_version(script_runner):
    ret = script_runner.run(["cat-numbers", "--version"], print_result=False)
    assert ret.success
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_help(script_runner):
    ret = script_runner.run(["cat-numbers", "--help"], print_result=False)
    assert ret.success
    assert "List the names of tables" in ret.stdout
    assert "Names of sheet" in ret.stdout
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_full_contents(script_runner):
    ref = ""
    for row in ZZZ_TABLE_1_REF:
        ref += (
            f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_1: "
            + ",".join(["" if v is None else v for v in row])
            + "\n"
        )
    for row in ZZZ_TABLE_2_REF:
        ref += (
            f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_2: "
            + ",".join(["" if v is None else v for v in row])
            + "\n"
        )
    for row in XXX_TABLE_1_REF:
        ref += (
            f"{DOCUMENT}: ZZZ_Sheet_2: XXX_Table_1: "
            + ",".join(["" if v is None else v for v in row])
            + "\n"
        )
    ret = script_runner.run(["cat-numbers", DOCUMENT], print_result=False)
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_brief_contents(script_runner):
    ref = ""
    for row in ZZZ_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in ZZZ_TABLE_2_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in XXX_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run(["cat-numbers", "--brief", DOCUMENT], print_result=False)
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_select_sheet(script_runner):
    ref = ""
    for row in ZZZ_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in ZZZ_TABLE_2_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run(
        ["cat-numbers", "--sheet", "ZZZ_Sheet_1", "--brief", DOCUMENT],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_select_table(script_runner):
    ref = ""
    for row in XXX_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run(
        ["cat-numbers", "--table", "XXX_Table_1", "--brief", DOCUMENT],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_list_sheets(script_runner):
    ret = script_runner.run(["cat-numbers", "-S", DOCUMENT], print_result=False)
    assert ret.success
    assert ret.stdout == (f"{DOCUMENT}: ZZZ_Sheet_1\n" "tests/data/test-1.numbers: ZZZ_Sheet_2\n")
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_list_tables(script_runner):
    ret = script_runner.run(["cat-numbers", "-T", DOCUMENT], print_result=False)
    assert ret.success
    assert ret.stdout == (
        f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_1\n"
        f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_2\n"
        f"{DOCUMENT}: ZZZ_Sheet_2: XXX_Table_1\n"
    )
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_without_formulas(script_runner):
    ret = script_runner.run(
        ["cat-numbers", "-b", "-t", "Table 2", "tests/data/test-10.numbers"],
        print_result=False,
    )
    assert ret.success
    rows = ret.stdout.strip().split("\n")
    assert rows == [
        "XXX_1,XXX_1XXX_2XXX_3",
        "XXX_2,10.0",
        "XXX_3,X",
        "XXX_4,XX",
        "XXX_5,_5",
        "XXX_6,4.0",
        "XXX_7,#REF!",
        "XXX_8,XXX_1",
        "0.25,0.5",
        "2.0,smaller",
        "10.0,larger",
    ]

    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_with_formulas(script_runner):
    ret = script_runner.run(
        [
            "cat-numbers",
            "-b",
            "-t",
            "Table 2",
            "--formulas",
            "tests/data/test-10.numbers",
        ],
        print_result=False,
    )
    assert ret.success
    rows = ret.stdout.strip().split("\n")
    assert rows == [
        "XXX_1,A1&A2&A3",
        "XXX_2,LEN(A2)+LEN(A3)",
        'XXX_3,"LEFT(A3,1)"',
        'XXX_4,"MID(A4,2,2)"',
        'XXX_5,"RIGHT(A5,2)"',
        'XXX_6,"FIND(""_"",A6)"',
        'XXX_7,"FIND(""YYY"",A7)"',
        'XXX_8,"IF(FIND(""_"",A8)>2,A1,A2)"',
        "0.25,100×(A9×2)%",
        '2.0,"IF(A10<5,""smaller"",""larger"")"',
        '10.0,"IF(A11≤5,""smaller"",""larger"")"',
    ]


@pytest.mark.script_launch_mode("subprocess")
def test_duration_formatting(script_runner):
    ret = script_runner.run(
        ["cat-numbers", "-b", "--formatting", "tests/data/duration_112.numbers"],
        print_result=False,
    )
    assert ret.success
    rows = ret.stdout.strip().splitlines()
    csv_reader = csv.reader(rows)
    for row in csv_reader:
        if row[13] != "Check" and row[13] is not None:
            check.equal(row[6], row[13])


@pytest.mark.script_launch_mode("subprocess")
def test_date_formatting(script_runner):
    ret = script_runner.run(
        ["cat-numbers", "-b", "--formatting", "tests/data/date_formats.numbers"],
        print_result=False,
    )
    assert ret.success
    rows = ret.stdout.strip().splitlines()
    csv_reader = csv.reader(rows)
    for row in csv_reader:
        if row[7] != "Check" and row[7] is not None:
            check.equal(row[6], row[7])


def test_debug(script_runner):
    ret = script_runner.run(
        ["cat-numbers", "--experimental", "--debug", "tests/data/test-1.numbers"],
        print_result=False,
    )
    assert ret.success
    rows = ret.stderr.strip().splitlines()
    assert rows[0] == "DEBUG:numbers_parser.experimental:Experimental features on"

    _enable_experimental_features(False)


@pytest.mark.script_launch_mode("subprocess")
def test_main(script_runner):
    ret = script_runner.run(
        ["python3", "-m", "numbers_parser._cat_numbers", "--help"], print_result=False
    )
    assert ret.success
    assert "List the names of tables" in ret.stdout
    assert "Names of sheet" in ret.stdout
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_errors(script_runner):
    ret = script_runner.run(["cat-numbers", "tests/data/corrupted.numbers"], print_result=False)
    assert not ret.success
    assert "Index/Metadata.iwa: invalid" in ret.stderr
    assert ret.stdout == ""

    ret = script_runner.run(["cat-numbers", "tests/data/badzipfile.numbers"], print_result=False)
    assert not ret.success
    assert "tests/data/badzipfile.numbers: Invalid Numbers file" in ret.stderr
    assert ret.stdout == ""

    ret = script_runner.run(["cat-numbers", "tests/data/badindexzip.numbers"], print_result=False)
    assert not ret.success
    assert "tests/data/badindexzip.numbers: Invalid Numbers file" in ret.stderr
    assert ret.stdout == ""

    ret = script_runner.run(["cat-numbers", "tests/data/badindexzip2.numbers"], print_result=False)
    assert not ret.success
    assert "tests/data/badindexzip2.numbers: Invalid Numbers file" in ret.stderr
    assert ret.stdout == ""

    ret = script_runner.run(["cat-numbers", "invalid.numbers"], print_result=False)
    assert not ret.success
    assert "invalid.numbers: No such file or directory" in ret.stderr
    assert ret.stdout == ""
