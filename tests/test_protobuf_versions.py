import os
import pytest
import re
import subprocess

from numbers_parser import _get_version

SOURCE_DIR = os.getcwd()
PROTOBUF_VERSIONS = ["4.21.1", "3.20.1"]


@pytest.mark.experimental
def test_protobuf_versions(virtualenv):
    result = subprocess.check_output(["python3", "setup.py", "bdist_wheel"], text=True)
    match = re.search("creating r'(.*?[.]whl)'", result)
    wheel = None
    if match:
        wheel = match.group(1)
    assert wheel is not None

    virtualenv.run(["pip", "install", f"{SOURCE_DIR}/{wheel}"])

    for protobuf_version in PROTOBUF_VERSIONS:
        virtualenv.install_package(
            "protobuf", version=protobuf_version, installer="pip"
        )
        pip_output = virtualenv.run(["pip", "show", "protobuf"], capture=True)
        match = re.search(f"Version: {protobuf_version}", pip_output)
        assert match is not None

        version = virtualenv.run("cat-numbers --version", capture=True).strip()
        assert version == _get_version()
