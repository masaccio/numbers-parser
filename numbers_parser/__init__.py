"""Extract data from Apple Numbers spreadsheets."""
__author__ = "Jon Connell"

import numbers_parser.macos_app_version
from numbers_parser.document import Document

__major_version__ = 1
__patch_version__ = 0
__supported_numbers_version__ = numbers_parser.macos_app_version.MacOSAppVersion(
    "10.3.9", "7029.9.8", "1A22"
)
__version_tuple__ = (
    __major_version__,
    __supported_numbers_version__.major,
    __supported_numbers_version__.minor,
    __patch_version__,
)
__version__ = ".".join([str(x) for x in __version_tuple__])

__email__ = "github@figsandfudge.com"
__description__ = 'A tool for reading Apple Numbers spreadsheets.'
__url__ = "https://github.com/masaccio/numbers-parser"
__new_issue_url__ = "https://github.com/masaccio/numbers-parser/issues/new"
