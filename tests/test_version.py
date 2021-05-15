import pytest

import numbers_parser

numbers_parser._SUPPORTED_NUMBERS_VERSIONS = ["0.0"]

def test_version(capsys):
    with pytest.warns(UserWarning) as record:
        numbers_parser._check_installed_numbers_version()
    assert "not tested with this version" in record[0].message.args[0]
