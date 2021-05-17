import pytest

from numbers_parser import __version__

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


def test_no_documents(script_runner):
    ret = script_runner.run("cat-numbers", "--brief", print_result=False)
    assert ret.success
    assert "usage: cat-numbers" in ret.stdout
    assert ret.stderr == ""


def test_version(script_runner):
    ret = script_runner.run("cat-numbers", "--version", print_result=False)
    assert ret.success
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""


def test_help(script_runner):
    ret = script_runner.run("cat-numbers", "--help", print_result=False)
    assert ret.success
    assert "List the names of tables" in ret.stdout
    assert "Names of sheet" in ret.stdout
    assert ret.stderr == ""


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
    ret = script_runner.run("cat-numbers", DOCUMENT, print_result=False)
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


def test_brief_contents(script_runner):
    ref = ""
    for row in ZZZ_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in ZZZ_TABLE_2_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in XXX_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run("cat-numbers", "--brief", DOCUMENT, print_result=False)
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


def test_select_sheet(script_runner):
    ref = ""
    for row in ZZZ_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    for row in ZZZ_TABLE_2_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run(
        "cat-numbers",
        "--sheet",
        "ZZZ_Sheet_1",
        "--brief",
        DOCUMENT,
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


def test_select_table(script_runner):
    ref = ""
    for row in XXX_TABLE_1_REF:
        ref += ",".join(["" if v is None else v for v in row]) + "\n"
    ret = script_runner.run(
        "cat-numbers",
        "--table",
        "XXX_Table_1",
        "--brief",
        DOCUMENT,
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ref
    assert ret.stderr == ""


def test_list_sheets(script_runner):
    ret = script_runner.run("cat-numbers", "-S", DOCUMENT, print_result=False)
    assert ret.success
    assert ret.stdout == (f"{DOCUMENT}: ZZZ_Sheet_1\n" "tests/data/test-1.numbers: ZZZ_Sheet_2\n")
    assert ret.stderr == ""


def test_list_tables(script_runner):
    ret = script_runner.run("cat-numbers", "-T", DOCUMENT, print_result=False)
    assert ret.success
    assert ret.stdout == (
        f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_1\n"
        f"{DOCUMENT}: ZZZ_Sheet_1: ZZZ_Table_2\n"
        f"{DOCUMENT}: ZZZ_Sheet_2: XXX_Table_1\n"
    )
    assert ret.stderr == ""
