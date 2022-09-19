import os

from pkg_resources import resource_filename
from datetime import datetime

# New document defaults
DEFAULT_DOCUMENT = resource_filename(__name__, os.path.join("data", "empty.numbers"))
DEFAULT_ROW_COUNT = 10
DEFAULT_COLUMN_COUNT = 5
DEFAULT_TABLE_OFFSET = 80.0

# Numbers limits
MAX_TILE_SIZE = 256
MAX_ROW_COUNT = 1000000
MAX_COL_COUNT = 1000

# Root object IDs
DOCUMENT_ID = 1
PACKAGE_ID = 2

# System constants
EPOCH = datetime(2001, 1, 1)
