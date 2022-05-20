import os

from pkg_resources import resource_filename
from datetime import datetime

DEFAULT_DOCUMENT = resource_filename(__name__, os.path.join("data", "empty.numbers"))
DEFAULT_ROW_COUNT = 10
DEFAULT_COLUMN_COUNT = 5
EPOCH = datetime(2001, 1, 1)
MAX_TILE_SIZE = 256
DOCUMENT_ID = 1
PACKAGE_ID = 2
