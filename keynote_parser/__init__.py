# Copyright 2020 Peter Sobot
"""Unpack and repack Apple Keyote files."""
__author__ = "Peter Sobot"

import keynote_parser.macos_app_version

__supported_keynote_version__ = keynote_parser.macos_app_version.MacOSAppVersion(
    "10.0", "6748", "1A171"
)
__version_tuple__ = (1, __supported_keynote_version__.major, 0)
__version__ = ".".join([str(x) for x in __version_tuple__])

__email__ = "github@petersobot.com"
__description__ = 'A tool for manipulating Apple Keynote presentation files.'
__url__ = "https://github.com/psobot/keynote-parser"
__new_issue_url__ = "https://github.com/psobot/keynote-parser/issues/new"
__command_line_invocation__ = False
