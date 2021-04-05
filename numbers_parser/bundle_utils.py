import os
import warnings
import plistlib
from numbers_parser import (
    __version__,
    __supported_numbers_version__,
)
from numbers_parser.macos_app_version import MacOSAppVersion

DEFAULT_NUMBERS_INSTALL_PATH = "/Applications/Numbers.app"
VERSION_PLIST_PATH = "Contents/version.plist"


def get_installed_numbers_version():
    try:
        fp = open(os.path.join(DEFAULT_NUMBERS_INSTALL_PATH, VERSION_PLIST_PATH), "rb")
    except IOError:
        return None
    version_dict = plistlib.load(fp)
    return MacOSAppVersion(
        version_dict["CFBundleShortVersionString"],
        version_dict["CFBundleVersion"],
        version_dict["ProductBuildVersion"],
    )


class NumbersVersionWarning(UserWarning):
    def __init__(self, installed_numbers_version):
        issue_title = "Please add support for Numbers %s" % installed_numbers_version
        super(UserWarning, self).__init__(
            (
                "This numbers_parser (%s) was not built for the installed Numbers.app:\n"
                "\tnumbers_parser version:    %s\n"
                "\tsupported Numbers version: %s\n"
                "\tinstalled Numbers version: %s\n"
            )
            % (
                __version__,
                __version__,
                __supported_numbers_version__,
                installed_numbers_version,
            )
        )


DID_WARN = False


def warn_once_on_newer_numbers():
    global DID_WARN
    if DID_WARN:
        return False

    installed_numbers_version = get_installed_numbers_version()
    if not installed_numbers_version:
        return False

    if __supported_numbers_version__ < installed_numbers_version:
        print(NumbersVersionWarning(installed_numbers_version))
        DID_WARN = True

    return DID_WARN

warn_once_on_newer_numbers()
