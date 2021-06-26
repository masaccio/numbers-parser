from numbers_parser.containers import NumbersError
from datetime import datetime, timedelta


class Cell:
    def __init__(self, row: int, col: int, value=None):
        self._value = value
        self.row = row
        self.col = col
        pass


class NumberCell(Cell):
    @property
    def value(self) -> int:
        return self._value


class TextCell(Cell):
    @property
    def value(self) -> str:
        return self._value


class EmptyCell(Cell):
    @property
    def value(self):
        return None


class BoolCell(Cell):
    @property
    def value(self) -> bool:
        return self._value


class DateCell(Cell):
    @property
    def value(self) -> timedelta:
        return self._value


class DurationCell(Cell):
    @property
    def value(self) -> timedelta:
        return self._value


class FormulaCell(Cell):
    @property
    def value(self):
        return None


class ErrorCell(Cell):
    @property
    def value(self):
        return None


class MergedCell:
    def __init__(self, row_start: int, row_end: int, col_start: int, col_end: int):
        self.value = None
        self.row_start = row_start
        self.row_end = row_end
        self.col_start = col_start
        self.col_end = col_end
        pass
