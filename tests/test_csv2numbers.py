"""Tests for CSV conversion."""

import shutil
from pathlib import Path

import pytest

from numbers_parser import Document, _get_version
from numbers_parser._csv2numbers import Transformer


@pytest.mark.script_launch_mode("subprocess")
def test_help(script_runner) -> None:
    """Test conversion with no transforms."""
    ret = script_runner.run(["csv2numbers"], print_result=False)
    assert not ret.success
    assert "At least one CSV file is required" in ret.stderr

    ret = script_runner.run(["csv2numbers", "--help"], print_result=False)
    assert ret.success
    assert "usage: csv2numbers" in ret.stdout


@pytest.mark.script_launch_mode("subprocess")
def test_version(script_runner) -> None:
    """Test Version number."""
    ret = script_runner.run(["csv2numbers", "-V"], print_result=False)
    assert ret.stderr == ""
    assert ret.stdout.strip() == _get_version()


def test_defaults(script_runner, tmp_path) -> None:
    """Test conversion with no options."""
    csv_path = str(tmp_path / "format-1.csv")
    shutil.copy("tests/data/format-1.csv", csv_path)

    ret = script_runner.run(["csv2numbers", csv_path], print_result=False)
    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")

    assert numbers_path.exists()
    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    assert table.cell(3, 1).value == "GROCERY STORE        LONDON"


@pytest.mark.script_launch_mode("inprocess")
def test_defaults_no_header(script_runner, tmp_path) -> None:
    """Test conversion with no options."""
    csv_path = str(tmp_path / "format-1.csv")
    shutil.copy("tests/data/format-1.csv", csv_path)

    ret = script_runner.run(["csv2numbers", "--no-header", csv_path], print_result=False)
    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")

    assert numbers_path.exists()
    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    assert table.cell(3, 1).value == "GROCERY STORE        LONDON"


@pytest.mark.script_launch_mode("subprocess")
def test_errors(script_runner) -> None:
    """Test error detection in command line."""
    ret = script_runner.run(
        ["csv2numbers", "--delete=XX", "tests/data/format-1.csv"],
        print_result=False,
    )
    assert "'XX': cannot delete" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", "--transform=XX=POS:YY", "tests/data/format-1.csv"],
        print_result=False,
    )
    assert "transform failed: column(s) do not exist in CSV" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", "--transform=XX=FUNC:Account", "tests/data/format-1.csv"],
        print_result=False,
    )
    assert "'FUNC': invalid transformation" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", '--rename=Foo«bar,"a,b'],
        print_result=False,
    )
    assert "malformed CSV string" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", "--rename=foo"],
        print_result=False,
    )
    assert "column rename maps must be formatted" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", '--delete=Foo«bar,"a,b'],
        print_result=False,
    )
    assert "can't parse argument" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", '--transform=Foo«bar,"a,b'],
        print_result=False,
    )
    assert "malformed CSV string" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", "--transform=foo"],
        print_result=False,
    )
    assert "invalid transformation format" in ret.stderr

    ret = script_runner.run(
        ["csv2numbers", "not-exists.csv"],
        print_result=False,
    )
    assert "not-exists.csv: file not found" in ret.stderr

    ret = script_runner.run(
        [
            "csv2numbers",
            "--transform=Category=LOOKUP:Description",
            "tests/data/matches.csv",
        ],
        print_result=False,
    )
    assert "LOOKUP must have exactly 2 arguments" in ret.stderr

    ret = script_runner.run(
        [
            "csv2numbers",
            "--transform=Category=LOOKUP:XX;YY",
            "tests/data/matches.csv",
        ],
        print_result=False,
    )
    assert "no such file or directory" in ret.stderr

    ret = script_runner.run(
        [
            "csv2numbers",
            "--transform=Category=LOOKUP:XX;tests/data/mapping.numbers",
            "tests/data/matches.csv",
        ],
        print_result=False,
    )
    assert "transform failed: column(s) do not exist in CSV" in ret.stderr

    ret = script_runner.run(
        [
            "csv2numbers",
            "--transform=Category=LOOKUP:Description;tests/data/corrupted.numbers",
            "tests/data/matches.csv",
        ],
        print_result=False,
    )
    assert "Index/Metadata.iwa: invalid IWA file Index/Metadata.iwa" in ret.stderr


@pytest.mark.script_launch_mode("subprocess")
def test_multifile(script_runner, tmp_path) -> None:
    """Test conversion with no options."""
    csv_path_1 = str(tmp_path / "format-1.csv")
    shutil.copy("tests/data/format-1.csv", csv_path_1)
    numbers_path_1 = Path(csv_path_1).with_suffix(".numbers")
    csv_path_2 = str(tmp_path / "format-2.csv")
    shutil.copy("tests/data/format-2.csv", csv_path_2)
    numbers_path_2 = Path(csv_path_2).with_suffix(".numbers")

    ret = script_runner.run(
        ["csv2numbers", csv_path_1, csv_path_2, "-o", numbers_path_1, numbers_path_2],
        print_result=False,
    )
    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success

    assert numbers_path_1.exists()
    assert numbers_path_2.exists()

    ret = script_runner.run(
        [
            "csv2numbers",
            "tests/data/format-1.csv",
            "tests/data/format-2.csv",
            "--output",
            "invalid",
        ],
        print_result=False,
    )
    assert "numbers of input and output file names do not match" in ret.stderr


@pytest.mark.script_launch_mode("subprocess")
def test_parse_error(script_runner) -> None:
    """Test conversion with no transforms."""
    ret = script_runner.run(
        ["csv2numbers", "tests/data/error.csv"],
        print_result=False,
    )
    assert "tests/data/error.csv@2: unexpected end of data" in ret.stderr


@pytest.mark.script_launch_mode("inprocess")
def test_transforms_format_1(script_runner, tmp_path) -> None:
    """Test conversion with transformation."""
    csv_path = str(tmp_path / "format-1.csv")
    shutil.copy("tests/data/format-1.csv", csv_path)

    ret = script_runner.run(
        [
            "csv2numbers",
            "--whitespace",
            "--reverse",
            "--day-first",
            "--date=Date",
            "--delete=Card Member,Account #",
            csv_path,
        ],
        print_result=False,
    )
    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")
    assert numbers_path.exists()

    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    assert table.cell(1, 1).value == "FLOWERS INC. 202-5551234"
    assert str(table.cell(2, 0).value) == "2008-04-02 00:00:00"
    assert table.cell(6, 2).value == 30.99


@pytest.mark.script_launch_mode("inprocess")
def test_transforms_format_2(script_runner, tmp_path) -> None:
    """Test conversion with transformation."""
    csv_path = str(tmp_path / "format-2.csv")
    shutil.copy("tests/data/format-2.csv", csv_path)

    ret = script_runner.run(
        [
            "csv2numbers",
            "--whitespace",
            "--day-first",
            "--date=Date",
            "--transform=Paid In=POS:Amount,Withdrawn=NEG:Amount",
            "--delete=Amount,Balance",
            csv_path,
        ],
        print_result=False,
    )
    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")
    assert numbers_path.exists()

    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    assert table.cell(0, 2).value == "Paid In"
    assert table.cell(0, 3).value == "Withdrawn"
    assert table.cell(1, 3).value == 1.4
    assert table.cell(3, 2).value == 10.0
    assert str(table.cell(3, 0).value) == "2003-02-04 00:00:00"


@pytest.mark.script_launch_mode("inprocess")
def test_transforms_format_3(script_runner, tmp_path) -> None:
    """Test conversion with transformation."""
    csv_path = str(tmp_path / "format-3.csv")
    shutil.copy("tests/data/format-3.csv", csv_path)

    ret = script_runner.run(
        [
            "csv2numbers",
            "--delete=2,3,4,5",
            "--date=0",
            "--day-first",
            "--no-header",
            "--rename=0:Date,1:Transaction,6:Amount",
            "--transform=6=MERGE:5;6",
            "--whitespace",
            csv_path,
        ],
        print_result=False,
    )

    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")
    assert numbers_path.exists()

    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    assert table.cell(5, 1).value == "AutoShop.com"
    assert str(table.cell(7, 0).value) == "2023-09-26 00:00:00"
    assert table.cell(0, 1).value == "Transaction"
    assert table.cell(0, 2).value == "Amount"
    assert table.cell(7, 2).value == -1283.72


@pytest.mark.script_launch_mode("subprocess")
def test_transforms_lookup(script_runner, tmp_path) -> None:
    """Test conversion with transformation."""
    csv_path = str(tmp_path / "matches.csv")
    shutil.copy("tests/data/matches.csv", csv_path)

    ret = script_runner.run(
        [
            "csv2numbers",
            "--transform=Category=LOOKUP:Description;tests/data/mapping.numbers",
            csv_path,
        ],
        print_result=False,
    )

    assert ret.stdout == ""
    assert ret.stderr == ""
    assert ret.success
    numbers_path = Path(csv_path).with_suffix(".numbers")
    assert numbers_path.exists()

    doc = Document(str(numbers_path))
    table = doc.sheets[0].tables[0]
    categories = [table.cell(row_num, 3).value for row_num in range(table.num_rows)]
    assert categories == ["Category", "Groceries", "Fuel", "Clothes", "", "Flowers"]

    cls = Transformer("XX", "YY")
    with pytest.raises(NotImplementedError):
        cls.transform_row(None)
