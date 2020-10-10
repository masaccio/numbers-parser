# Copyright 2020 Peter Sobot
"""Unpack and repack Apple Keyote files."""
__author__ = "Peter Sobot"

import keynote_parser.macos_app_version

__major_version__ = 1
__patch_version__ = 0
__supported_keynote_version__ = keynote_parser.macos_app_version.MacOSAppVersion(
    "10.2", "7028.0.88", "1A122"
)
__version_tuple__ = (
    __major_version__,
    __supported_keynote_version__.major,
    __supported_keynote_version__.minor,
    __patch_version__,
)
__version__ = ".".join([str(x) for x in __version_tuple__])

__email__ = "github@petersobot.com"
__description__ = 'A tool for manipulating Apple Keynote presentation files.'
__url__ = "https://github.com/psobot/keynote-parser"
__new_issue_url__ = "https://github.com/psobot/keynote-parser/issues/new"
__command_line_invocation__ = False
