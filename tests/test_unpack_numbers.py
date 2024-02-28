import json
import sys

import magic
import pytest

from numbers_parser import __version__


def test_version(script_runner):
    ret = script_runner.run(["unpack-numbers", "--version"], print_result=False)
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""
    assert ret.success


def test_help(script_runner):
    ret = script_runner.run(["unpack-numbers", "--help"], print_result=False)
    assert "directory name to unpack into" in ret.stdout
    assert "document" in ret.stdout
    assert ret.stderr == ""
    assert ret.success

    ret = script_runner.run(["unpack-numbers"], print_result=False)
    assert "directory name to unpack into" in ret.stdout
    assert "document" in ret.stdout
    assert ret.stderr == ""
    assert ret.success


def test_multi_doc_error(script_runner):
    ret = script_runner.run(["unpack-numbers", "--output", "tmp", "foo", "bar"], print_result=False)
    assert not ret.success
    assert ret.stdout == ""
    assert "output directory only valid" in ret.stderr


def test_unpack_file(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        ["unpack-numbers", "--output", str(output_dir), "tests/data/test-1.numbers"],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ""
    assert (output_dir / "preview.jpg").exists()
    assert "JPEG image data" in magic.from_file(str(output_dir / "preview.jpg"))
    assert (output_dir / "Index/CalculationEngine.json").exists()
    assert (output_dir / "Index/Tables/DataList-954857.json").exists()
    with open(str(output_dir / "Index/Tables/DataList-954857.json")) as f:
        data = json.load(f)
    objects = data["chunks"][0]["archives"][0]["objects"]
    strings = [x["string"] for x in objects[0]["entries"]]
    assert len(strings) == 15
    assert "ZZZ_2_3" in strings
    assert "ZZZ_ROW_3" in strings


def test_unpack_dir(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        ["unpack-numbers", "--output", str(output_dir), "tests/data/test-5.numbers"],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ""
    assert (output_dir / "preview.jpg").exists()
    assert "JPEG image data" in magic.from_file(str(output_dir / "preview.jpg"))
    assert (output_dir / "Index/CalculationEngine.json").exists()
    assert (output_dir / "Index/Tables/DataList-875166.json").exists()
    with open(str(output_dir / "Index/Tables/DataList-875166.json")) as f:
        data = json.load(f)
    objects = data["chunks"][0]["archives"][0]["objects"]
    strings = [x["string"] for x in objects[0]["entries"]]
    assert len(strings) == 21
    assert "XXX_3_3" in strings
    assert "XXX_COL_3" in strings


@pytest.mark.script_launch_mode("inprocess")
def test_unpack_hex(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        [
            "unpack-numbers",
            "--hex-uuids",
            "--output",
            str(output_dir),
            "tests/data/test-5.numbers",
        ],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ""
    with open(str(output_dir / "Index/CalculationEngine.json")) as f:
        data = json.load(f)
    objects = data["chunks"][0]["archives"][1]["objects"][0]
    assert objects["base_owner_uid"] == "83aa364c-869c-498a-b749-dbddb35f99d7"
    objects = data["chunks"][0]["archives"][0]["objects"][0]
    assert (
        objects["dependency_tracker"]["formula_owner_info"][0]["formula_owner_id"]
        == "83aa364c-869c-498a-b749-dbddb35f99d7"
    )


def test_pretty_storage(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        [
            "unpack-numbers",
            "--pretty-storage",
            "--output",
            str(output_dir),
            "tests/data/test-5.numbers",
        ],
        print_result=False,
    )
    assert ret.stderr == ""
    assert ret.stdout == ""
    assert ret.success
    with open(str(output_dir / "Index/Tables/Tile-875165.json")) as f:
        data = json.load(f)
    objects = data["chunks"][0]["archives"][0]["objects"][0]
    assert objects["rowInfos"][0]["cell_offsets"] == "-1,0,24,48,72,96,[...]"
    if sys.version_info.minor >= 8:
        assert objects["rowInfos"][0]["cell_storage_buffer"][0:26] == "05:03:00:00:00:00:00:00:08"
    else:
        assert objects["rowInfos"][0]["cell_storage_buffer"][0:26] == "05030000000000000810020002"


def test_compact_json(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        [
            "unpack-numbers",
            "--compact-json",
            "--output",
            str(output_dir),
            "tests/data/test-5.numbers",
        ],
        print_result=False,
    )
    assert ret.success
    assert ret.stdout == ""

    with open(str(output_dir / "Index/CalculationEngine.json")) as f:
        data = f.read()
    assert (
        '"column": 0, "row": 1, "contains_a_formula": true, "edges": {"packed_edge_without_owner":'
        in data
    )


@pytest.mark.script_launch_mode("subprocess")
def test_main(script_runner):
    ret = script_runner.run(
        ["python3", "-m", "numbers_parser._unpack_numbers", "--help"],
        print_result=False,
    )
    assert ret.success
    assert "directory name to unpack into" in ret.stdout
    assert ret.stderr == ""


@pytest.mark.script_launch_mode("subprocess")
def test_corrupted(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        ["unpack-numbers", "--output", str(output_dir), "tests/data/corrupted.numbers"],
        print_result=False,
    )
    assert not ret.success
    assert "Index/Metadata.iwa: invalid" in ret.stderr
    assert ret.stdout == ""


def test_debug(script_runner, tmp_path):
    output_dir = tmp_path / "test"
    ret = script_runner.run(
        [
            "unpack-numbers",
            "--debug",
            "--output",
            str(output_dir),
            "tests/data/test-1.numbers",
        ],
        print_result=False,
    )
    assert ret.success
    rows = ret.stderr.strip().splitlines()
    assert "DEBUG:numbers_parser.iwork:open: filename=" in rows[0]
