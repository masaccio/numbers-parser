import pytest

from numbers_parser import __version__

DOCUMENT = "tests/data/test-1.numbers"

def test_version(script_runner):
    ret = script_runner.run("scripts/unpack-numbers", "--version", print_result=False)
    assert ret.success
    assert ret.stdout == __version__ + "\n"
    assert ret.stderr == ""


def test_help(script_runner):
    ret = script_runner.run("scripts/unpack-numbers", "--help", print_result=False)
    assert ret.success
    assert "directory name to unpack into" in ret.stdout
    assert "document" in ret.stdout
    assert ret.stderr == ""


def test_multi_doc_error(script_runner):
    ret = script_runner.run("scripts/unpack-numbers", "--output", "tmp", "foo", "bar", print_result=False)
    assert ret.success == False
    assert ret.stdout == ""
    assert "output directory only valid" in ret.stderr
