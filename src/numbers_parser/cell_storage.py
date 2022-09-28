from numbers_parser.exceptions import UnsupportedError
from numbers_parser.constants import EPOCH
from struct import unpack
from datetime import timedelta


class CellStorage:
    def __init__(self, buffer: bytes):
        version = buffer[0]
        if version != 5:  # pragma: no cover
            raise UnsupportedError(f"Cell storage version {version} is unsupported")

        self._buffer = buffer
        self._flags = unpack("<i", buffer[8:12])[0]
        self._current_offset = 12

        self.d128 = self.pop_buffer(16) if self._flags & 0x1 else None
        self.double = self.pop_buffer(8) if self._flags & 0x2 else None
        self.datetime = (
            EPOCH + timedelta(seconds=self.pop_buffer(8)) if self._flags & 0x4 else None
        )
        self.string_id = self.pop_buffer() if self._flags & 0x8 else None
        self.rich_id = self.pop_buffer() if self._flags & 0x10 else None
        self.cell_style_id = self.pop_buffer() if self._flags & 0x20 else None
        self.text_style_id = self.pop_buffer() if self._flags & 0x40 else None
        self.cond_style_id = self.pop_buffer() if self._flags & 0x80 else None
        self.cond_rule_style_id = self.pop_buffer() if self._flags & 0x100 else None
        self.formula_id = self.pop_buffer() if self._flags & 0x200 else None
        self.control_id = self.pop_buffer() if self._flags & 0x400 else None
        self.formula_error_id = self.pop_buffer() if self._flags & 0x800 else None
        self.suggest_id = self.pop_buffer() if self._flags & 0x1000 else None
        self.num_format_id = self.pop_buffer() if self._flags & 0x2000 else None
        self.currency_format_id = self.pop_buffer() if self._flags & 0x4000 else None
        self.date_format_id = self.pop_buffer() if self._flags & 0x8000 else None
        self.duration_format_id = self.pop_buffer() if self._flags & 0x10000 else None
        self.text_format_id = self.pop_buffer() if self._flags & 0x20000 else None
        self.bool_format_id = self.pop_buffer() if self._flags & 0x40000 else None
        self.comment_id = self.pop_buffer() if self._flags & 0x80000 else None
        self.import_warning_id = self.pop_buffer() if self._flags & 0x100000 else None

    def pop_buffer(self, size: int = 4):
        offset = self._current_offset
        self._current_offset += size
        if size == 16:
            return unpack_decimal128(self._buffer[offset : offset + 16])
        elif size == 8:
            return unpack("<d", self._buffer[offset : offset + 8])[0]
        else:
            return unpack("<i", self._buffer[offset : offset + 4])[0]


def unpack_decimal128(buffer: bytearray) -> float:
    exp = (((buffer[15] & 0x7F) << 7) | (buffer[14] >> 1)) - 0x1820
    mantissa = buffer[14] & 1
    for i in range(13, -1, -1):
        mantissa = mantissa * 256 + buffer[i]
    if buffer[15] & 0x80:
        mantissa = -mantissa
    value = mantissa * 10**exp
    return float(value)
