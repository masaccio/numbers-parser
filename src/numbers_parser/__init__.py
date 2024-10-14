"""Parse and extract data from Apple Numbers spreadsheets."""

import importlib.metadata
import os
import plistlib
import re
import warnings

with warnings.catch_warnings():
    # Protobuf
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    from numbers_parser.cell import *  # noqa: F403
    from numbers_parser.constants import *  # noqa: F403
    from numbers_parser.document import *  # noqa: F403
    from numbers_parser.exceptions import *  # noqa: F403

__version__ = importlib.metadata.version("numbers-parser")

_DEFAULT_NUMBERS_INSTALL_PATH = "/Applications/Numbers.app"
_VERSION_PLIST_PATH = "Contents/version.plist"


def _get_version() -> str:
    return __version__


def _check_installed_numbers_version() -> str:
    try:
        fp = open(os.path.join(_DEFAULT_NUMBERS_INSTALL_PATH, _VERSION_PLIST_PATH), "rb")
    except OSError:
        return None
    version_dict = plistlib.load(fp)
    installed_version = version_dict["CFBundleShortVersionString"]

    from numbers_parser.constants import SUPPORTED_NUMBERS_VERSIONS

    installed_version = re.sub(r"(\d+)\.(\d+)\.\d+", r"\1.\2", installed_version)
    if installed_version not in SUPPORTED_NUMBERS_VERSIONS:
        warnings.warn(
            f"Numbers version {installed_version} not tested with this version",
            stacklevel=2,
        )
    fp.close()
    return installed_version


_ = _check_installed_numbers_version()
