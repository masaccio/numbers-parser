import os

from pkg_resources import resource_filename
from datetime import datetime

# New document defaults
DEFAULT_DOCUMENT = resource_filename(__name__, os.path.join("data", "empty.numbers"))
DEFAULT_COLUMN_COUNT = 5
DEFAULT_COLUMN_WIDTH = 98.0
DEFAULT_PRE_BNC_BYTES = "🤠".encode("utf-8")  # Yes, really!
DEFAULT_ROW_COUNT = 10
DEFAULT_ROW_HEIGHT = 20.0
DEFAULT_TABLE_OFFSET = 80.0
DEFAULT_TILE_SIZE = 256


# Numbers limits
MAX_TILE_SIZE = 256
MAX_ROW_COUNT = 1000000
MAX_COL_COUNT = 1000
MAX_HEADER_COUNT = 5

# Root object IDs
DOCUMENT_ID = 1
PACKAGE_ID = 2

# System constants
EPOCH = datetime(2001, 1, 1)
