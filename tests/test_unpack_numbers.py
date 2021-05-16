import pytest

from numbers_parser import __version__

DOCUMENT = "tests/data/test-1.numbers"

def run_script(script_runner, *args):
    import pdb
    if script_runner.launch_mode == "inprocess":
        command = "scripts/unpack-numbers"
    else:
        command = "python3"
        args = tuple(["scripts/unpack-numbers"]) + args
    ret = script_runner.run(command, *args, print_result=False)
    return ret


def test_version(script_runner):
    ret = run_script(script_runner, "--version")
    assert ret.success
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""


def test_help(script_runner):
    ret = run_script(script_runner, "--help")
    assert ret.success
    assert "directory name to unpack into" in ret.stdout
    assert "document" in ret.stdout
    assert ret.stderr == ""


def test_multi_doc_error(script_runner):
    ret = run_script(script_runner, "--output", "tmp", "foo", "bar")
    assert ret.success == False
    assert ret.stdout == ""
    assert "output directory only valid" in ret.stderr


