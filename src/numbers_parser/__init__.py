import os
import plistlib
import warnings

from numbers_parser.document import Document

DEFAULT_NUMBERS_INSTALL_PATH = "/Applications/Numbers.app"
VERSION_PLIST_PATH = "Contents/version.plist"
SUPPORTED_NUMBERS_VERSIONS = [
    "10.3",
    "11.0",
]

#  Don't print the source line
formatwarning_old = warnings.formatwarning
warnings.formatwarning = lambda message, category, filename, lineno, line=None: formatwarning_old(
    message, category, filename, lineno, line=""
)


def check_installed_numbers_version():
    try:
        fp = open(os.path.join(DEFAULT_NUMBERS_INSTALL_PATH, VERSION_PLIST_PATH), "rb")
    except IOError:
        return None
    version_dict = plistlib.load(fp)
    installed_version = version_dict["CFBundleShortVersionString"]
    if installed_version not in SUPPORTED_NUMBERS_VERSIONS:
        warnings.warn(f"Numbers version {installed_version} not tested with this version")


check_installed_numbers_version()
