import json
import imghdr
import pytest

from numbers_parser import __version__


def test_version(script_runner):
    ret = script_runner.run("unpack-numbers", "--version", print_result=False)
    assert ret.success
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""


def test_help(script_runner):
    ret = script_runner.run("unpack-numbers", "--help", print_result=False)
    assert ret.success
    assert "directory name to unpack into" in ret.stdout
    assert "document" in ret.stdout
    assert ret.stderr == ""


def test_multi_doc_error(script_runner):
    ret = script_runner.run("unpack-numbers", "--output", "tmp", "foo", "bar", print_result=False)
    assert ret.success == False
    assert ret.stdout == ""
    assert "output directory only valid" in ret.stderr


def test_create_file(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        "unpack-numbers",
        "--output",
        str(output_dir),
        "tests/data/test-1.numbers",
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ""
    assert (output_dir / "preview.jpg").exists()
    assert imghdr.what(str(output_dir / "preview.jpg")) == "jpeg"
    assert (output_dir / "Index/CalculationEngine.txt").exists()
    assert (output_dir / "Index/Tables/DataList-954857.txt").exists()
    with open(str(output_dir / "Index/Tables/DataList-954857.txt")) as f:
        data = json.load(f)
    objects = data["chunks"][0]["archives"][0]["objects"]
    strings = [x["string"] for x in objects[0]["entries"]]
    assert len(strings) == 15
    assert "ZZZ_2_3" in strings
    assert "ZZZ_ROW_3" in strings
