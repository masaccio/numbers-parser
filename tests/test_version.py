import os
import pytest
import numbers_parser

numbers_parser._SUPPORTED_NUMBERS_VERSIONS = ["0.0"]

def test_version(capsys):
    if os.path.isfile(numbers_parser._DEFAULT_NUMBERS_INSTALL_PATH):
        with pytest.warns(UserWarning) as record:
            r = numbers_parser._check_installed_numbers_version()
        assert r is None or "not tested with this version" in record[0].message.args[0]
