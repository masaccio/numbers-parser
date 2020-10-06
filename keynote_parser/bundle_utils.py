import os
import warnings
import plistlib
import urllib
from colorama import init as colorama_init
from colorama import Fore
from keynote_parser import (
    __version__,
    __supported_keynote_version__,
    __new_issue_url__,
    __command_line_invocation__,
)
from keynote_parser.macos_app_version import MacOSAppVersion

DEFAULT_KEYNOTE_INSTALL_PATH = '/Applications/Keynote.app'
VERSION_PLIST_PATH = 'Contents/version.plist'


colorama_init()


def get_installed_keynote_version():
    try:
        fp = open(os.path.join(DEFAULT_KEYNOTE_INSTALL_PATH, VERSION_PLIST_PATH), 'rb')
    except IOError:
        return None
    version_dict = plistlib.load(fp)
    return MacOSAppVersion(
        version_dict['CFBundleShortVersionString'],
        version_dict['CFBundleVersion'],
        version_dict['ProductBuildVersion'],
    )


class KeynoteVersionWarning(UserWarning):
    def __init__(self, installed_keynote_version):
        issue_title = "Please add support for Keynote %s" % installed_keynote_version
        new_issue_url = __new_issue_url__ + "?" + urllib.parse.urlencode({"title": issue_title})
        super(UserWarning, self).__init__(
            (
                "KeynoteVersionWarning: " + Fore.RESET
                + "This version of keynote_parser (%s) was not built with support "
                "for\nthe currently installed version of Keynote.\n"
                "\tkeynote_parser version:    %s\n"
                "\tsupported Keynote version: %s\n"
                "\tinstalled Keynote version: %s\n"
                "\n"
                "Some presentation files may not be editable, or data corruption may occur.\n"
                "To get rid of this warning:\n"
                "\t- try %supgrading %skeynote-parser%s (pip install --upgrade keynote-parser)\n"
                "\t- %ssubmit an issue%s if no new version is available:\n"
                "\t\t%s\n"
                "\t- use the %swarnings%s module to suppress %sKeynoteVersionWarning%s:\n"
                "\t\tPYTHONWARNINGS=ignore:KeynoteVersionWarning ...\n"
            )
            % (
                __version__,
                __version__,
                __supported_keynote_version__,
                installed_keynote_version,
                Fore.CYAN,
                Fore.YELLOW,
                Fore.RESET,
                Fore.YELLOW,
                Fore.RESET,
                new_issue_url,
                Fore.CYAN,
                Fore.RESET,
                Fore.CYAN,
                Fore.RESET,
            )
        )


class CleanWarning(object):
    def custom_format_warning(self, message, *args):
        # Nasty hack - by putting the initial colour in the formatter,
        # the Warnings filter still lets users ignore this warning as
        # the message itself doesn't start with an ANSI escape sequence.
        return Fore.YELLOW + str(message) + "\n"

    def __enter__(self):
        self.old_formatwarning = warnings.formatwarning
        warnings.formatwarning = self.custom_format_warning

    def __exit__(self, *args, **kwargs):
        warnings.formatwarning = self.old_formatwarning


DID_WARN = False


def warn_once_on_newer_keynote():
    global DID_WARN
    if DID_WARN:
        return False

    installed_keynote_version = get_installed_keynote_version()
    if not installed_keynote_version:
        return False

    if __supported_keynote_version__ < installed_keynote_version:
        with CleanWarning():
            warnings.warn(KeynoteVersionWarning(installed_keynote_version))
        DID_WARN = True

    return DID_WARN


if not __command_line_invocation__:
    warn_once_on_newer_keynote()
