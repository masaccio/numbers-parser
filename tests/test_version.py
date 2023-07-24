import importlib
import mock
import pytest

from numbers_parser import _check_installed_numbers_version, _get_version

builtin_open = open


def test_version():
    version = _get_version()
    assert version.count(".") == 2


def mock_valid_plist(file, mode):
    if file.endswith("version.plist"):
        return builtin_open("tests/data/numbers-version-13.0.plist", mode)
    else:
        return builtin_open(file, mode)


def mock_newer_plist(file, mode):
    if file.endswith("version.plist"):
        return builtin_open("tests/data/numbers-version-99.0.plist", mode)
    else:
        return builtin_open(file, mode)


def mock_invalid_plist(file, mode):
    if file.endswith("version.plist"):
        return builtin_open("tests/data/XXXX.plist", mode)
    else:
        return builtin_open(file, mode)


def test_numbers_version_check():
    with mock.patch("builtins.open", side_effect=mock_valid_plist):
        import numbers_parser

        importlib.reload(numbers_parser)
        assert _check_installed_numbers_version() == "13.0"

    with mock.patch("builtins.open", side_effect=mock_invalid_plist):
        import numbers_parser

        importlib.reload(numbers_parser)
        assert _check_installed_numbers_version() is None

    with mock.patch("builtins.open", side_effect=mock_newer_plist):
        with pytest.warns(match="Numbers version 99.0 not tested with this version") as record:
            import numbers_parser

            importlib.reload(numbers_parser)
        assert len(record) == 1
