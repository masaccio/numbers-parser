"""Command-line utility to convert CSV files to Apple Numbers spreadsheets."""

# ruff: noqa: T201

from __future__ import annotations

import argparse
import contextlib
import csv
import re
import sys
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, cast

from dateutil.parser import parse  # type: ignore[import-untyped]

from numbers_parser import Document, NumbersError, _get_version, xl_rowcol_to_cell


@dataclass
class Converter:
    """Convert CSV data and formatting into a Numbers document."""

    input_filename: str | None = None
    output_filename: str | None = None
    date_columns: list | None = None
    day_first: bool = False
    no_header: bool = False
    reverse: bool = False
    whitespace: bool | None = None
    encoding: str = "utf-8"
    password: str | None = None
    header: list[Any] | None = field(init=False, default=None)
    data: list[Any] = field(init=False, default_factory=list)

    def __post_init__(self) -> None:
        """Parse CSV file with Pandas and return a dataframe."""
        self._read_csv()
        self._transform_data()

    def _read_csv(self) -> None:
        """Read a CSV file into the data array."""
        try:
            dialect = csv.excel
            dialect.strict = True
            lineno = 1
            with open(str(self.input_filename), encoding=self.encoding) as csvfile:
                csvreader = csv.reader(csvfile, dialect=dialect)
                if self.no_header:
                    self.header = None
                else:
                    self.header = [str(col or "") for col in next(csvreader)]
                    lineno += 1
                self.data = []
                for row in csvreader:
                    self.data.append([str(col or "") for col in row])
                    lineno += 1
        except FileNotFoundError as e:
            msg = f"{self.input_filename}: file not found"
            raise RuntimeError(msg) from e
        except csv.Error as e:
            msg = f"{self.input_filename}@{lineno}: {e.args[0]}"
            raise RuntimeError(msg) from e

    def _parse_date(self, x: str) -> datetime:
        """Parse a date string and return a datetime."""
        return parse(x, dayfirst=self.day_first).replace(tzinfo=timezone.utc)

    def _transform_data(self) -> None:
        """Apply type transformations to the data based in current configuration."""
        if self.no_header:
            self.header = list(range(len(self.data[0])))
        self.header = cast("list[str]", self.header)
        self.data = [dict(dict(zip(self.header, row, strict=False)).items()) for row in self.data]
        if self.reverse:
            self.data = list(reversed(self.data))
        is_date_column = {}
        if self.date_columns is not None and self.header is not None:
            date_columns = [self._column_key(col) for col in self.date_columns]
            is_date_column = {x: x in date_columns for x in self.header}
        for row in self.data:
            for k, v in row.items():
                if self.whitespace:
                    row[k] = re.sub(r"\s+", " ", v.strip())
                if self.date_columns is not None and is_date_column.get(k):
                    row[k] = self._parse_date(v)
                else:
                    # Attempt to coerce value into float
                    with contextlib.suppress(ValueError):
                        row[k] = float(v.replace(",", ""))

    def _column_key(self, column: str | int) -> str | int:
        """Resolve an index against named headers; retain keys for headerless CSVs."""
        header = self.header or []
        if isinstance(column, int) and column not in header and 0 <= column < len(header):
            return header[column]
        return column

    def rename_columns(self, mapper: dict | None) -> None:
        """Rename columns using column map."""
        if mapper is None:
            return
        mapper = {self._column_key(key): value for key, value in mapper.items()}
        self.no_header = False
        self.header = cast("list[str]", self.header)
        self.header = [mapper.get(x, x) for x in self.header]
        self.data = [
            {mapper.get(key, key): value for key, value in row.items()} for row in self.data
        ]

    def delete_columns(self, columns: list | None) -> None:
        """Delete columns from the data."""
        if columns is None:
            return

        self.header = cast("list[str]", self.header)
        columns = [self._column_key(col) for col in columns]
        missing_columns = list(set(columns) - set(self.header))
        if len(missing_columns) > 0:
            msg = "'" + "', '".join([str(x) for x in missing_columns]) + "'"
            msg += ": cannot delete: column(s) do not exist in CSV"
            raise RuntimeError(msg) from None

        for col in columns:
            self.header.remove(col)
        for row in self.data:
            for col in columns:
                del row[col]

    def transform_columns(self, columns: list[Transformer] | None) -> None:
        """Perform column transformations."""
        if columns is None:
            return
        for transform in columns:
            resolved = copy(transform)
            resolved.sources = [self._column_key(source) for source in transform.sources]
            resolved.dest = self._column_key(transform.dest)
            resolved.transform(self.header or [], self.data)
            if self.header is not None and resolved.dest not in self.header:
                self.header.append(resolved.dest)

    def save(self) -> None:
        """Write dataframe transctions to a Numbers file."""
        doc = Document(num_rows=2, num_cols=2, password=self.password)
        table = doc.sheets[0].tables[0]

        data: list[Any] = [] if self.no_header or self.header is None else [self.header]
        data += [list(row.values()) for row in self.data]

        for row_num, row in enumerate(data):
            for col_num, value in enumerate(row):
                table.write(row_num, col_num, value)
                if isinstance(value, datetime):
                    table.set_cell_formatting(
                        row_num,
                        col_num,
                        "datetime",
                        date_time_format="d MMM yyyy",
                    )

        doc.save(str(self.output_filename))


class Transformer:
    """Base class for column transformations."""

    def __init__(self, source: str, dest: str | int) -> None:
        """Initialize the instance state."""
        self.row_num = 0
        self.dest: str | int = int(dest) if isinstance(dest, str) and dest.isnumeric() else dest
        self.sources: list[str | int] = [
            int(x) if isinstance(x, str) and x.isnumeric() else x for x in str(source).split(";")
        ]

    def transform_row(self, row: dict) -> None:
        """Abstract base method for transforming rows using df.apply()."""
        raise NotImplementedError

    def transform(self, header: list[str | int], data: list[dict]) -> None:
        """Column transform to merge columns."""
        if not all(x in header for x in self.sources):
            missing = list(set(self.sources) - set(header))
            msg = "'" + "', '".join([str(x) for x in missing]) + "'"
            msg += ": transform failed: column(s) do not exist in CSV"
            raise RuntimeError(msg)
        self.row_num = len(header) - 1
        for row in data:
            self.transform_row(row)


class MergeTransformer(Transformer):
    """Transformer for column MERGE operations."""

    def transform_row(self, row: dict) -> None:
        """Merge data in a single row."""
        value = None
        for col, col_name in enumerate(self.sources):
            if row.get(col_name, "") != "":
                if value is not None:
                    ref = xl_rowcol_to_cell(self.row_num, col)
                    print(f"Warning: merge found multiple values for cell '{ref}'")
                else:
                    value = row[col_name]
        row[self.dest] = value if value is not None else ""


class NegTransformer(Transformer):
    """Transformer for column NEG operations."""

    def transform_row(self, row: dict) -> None:
        """Select negative values for a row."""
        value: float | str = ""
        for col in self.sources:
            if row.get(col, "") and not value and float(row[col]) < 0:
                value = abs(float(row[col]))
        row[self.dest] = value


class PosTransformer(Transformer):
    """Transformer for column POS operations."""

    def transform_row(self, row: dict) -> None:
        """Select positive values for a row."""
        value: float | str = ""
        for col in self.sources:
            if row.get(col, "") and not value and float(row[col]) > 0:
                value = float(row[col])
        row[self.dest] = value


class LookupTransformer(Transformer):
    """Transformer for column LOOKUP operations."""

    def __init__(self, source: str, dest: str | int) -> None:
        """Initialize the instance state."""
        super().__init__(source, dest)

        if len(self.sources) != 2:
            msg = f"'{self.sources}' LOOKUP must have exactly 2 arguments"
            raise RuntimeError(msg) from None

        source_val = self.sources[0]
        map_filename = str(self.sources[1])
        self.sources = [source_val]

        if not Path(map_filename).exists():
            msg = f"{map_filename}: no such file or directory"
            raise RuntimeError(msg) from None

        try:
            doc = Document(map_filename)
            table = doc.sheets[0].tables[0]
            self.lookup_map = {
                table.cell(row_num, 0).value: table.cell(row_num, 1).value  # type: ignore[attr-defined]
                for row_num in range(table.num_rows)
            }
        except NumbersError as e:
            msg = f"{map_filename}: {e!r}"
            raise RuntimeError(msg) from e

    def transform_row(self, row: dict) -> None:
        """Column transform to map values based on a lookup table."""
        matches = [
            {"value": v, "len": len(k)}
            for k, v in self.lookup_map.items()
            if re.search(str(k).lower(), str(row.get(self.sources[0], "")).lower())
        ]
        if len(matches) > 0:
            row[self.dest] = max(matches, key=lambda x: x["len"])["value"]
        else:
            row[self.dest] = ""


TRANSFORMERS = {
    "merge": MergeTransformer,
    "neg": NegTransformer,
    "pos": PosTransformer,
    "lookup": LookupTransformer,
}


def parse_columns(arg: str) -> list:
    """Parse a list of column names in Excel-compatible CSV format."""
    try:
        return [int(x) if x.isnumeric() else x for x in next(csv.reader([arg], strict=True))]
    except csv.Error as e:
        msg = f"'{arg}': can't parse argument"
        raise argparse.ArgumentTypeError(msg) from e


def parse_column_renames(arg: str) -> dict:
    """Parse a list of column renames in Excel-compatible CSV format."""
    mapper = {}
    try:
        for mapping in next(csv.reader([arg], strict=True)):
            if mapping.count(":") != 1:
                msg = f"'{mapping}': column rename maps must be formatted 'OLD:NEW'"
                raise argparse.ArgumentTypeError(msg)
            (old, new) = mapping.split(":")
            old_key = int(old) if old.isnumeric() else old
            mapper[old_key] = new
    except csv.Error as e:
        msg = f"'{arg}': malformed CSV string"
        raise argparse.ArgumentTypeError(msg) from e
    return mapper


def parse_column_transforms(arg: str) -> list[Transformer]:
    """Parse a list of column renames in Excel-compatible CSV format."""
    transforms = []
    try:
        for transform in next(csv.reader([arg], strict=True)):
            m = re.match(r"(.+)=(\w+):(.+)", transform)
            if not m:
                msg = f"'{transform}': invalid transformation format"
                raise argparse.ArgumentTypeError(msg)
            dest = m.group(1)
            func = m.group(2).lower()
            source = m.group(3)
            if func not in TRANSFORMERS:
                msg = f"'{m.group(2)}': invalid transformation"
                raise argparse.ArgumentTypeError(msg)
            transforms.append(TRANSFORMERS[func.lower()](source, dest))  # type: ignore[abstract]
    except csv.Error as e:
        msg = f"'{arg}': malformed CSV string"
        raise argparse.ArgumentTypeError(msg) from e
    return transforms


def command_line_parser() -> argparse.ArgumentParser:
    """Create a command-line argument parser and return parsed arguments."""
    parser = argparse.ArgumentParser()
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument(
        "--whitespace",
        required=False,
        action="store_true",
        help="strip whitespace from beginning and end of strings and "
        "collapse other whitespace into single space (default: false)",
    )
    parser.add_argument(
        "--reverse",
        required=False,
        action="store_true",
        help="reverse the order of the data rows (default: false)",
    )
    parser.add_argument(
        "--no-header",
        required=False,
        action="store_true",
        help="CSV file has no header row (default: false)",
    )
    parser.add_argument(
        "--day-first",
        required=False,
        action="store_true",
        help="dates are represented day first in the CSV file (default: false)",
    )
    parser.add_argument(
        "--encoding",
        required=False,
        default="utf-8",
        help="python-style text encoding of the CSV file (default: utf-8)",
    )
    parser.add_argument(
        "--date",
        metavar="COLUMNS",
        type=parse_columns,
        help="comma-separated list of column names/indexes to parse as dates",
    )
    parser.add_argument(
        "--rename",
        metavar="COLUMNS-MAP",
        type=parse_column_renames,
        help="comma-separated list of column names/indexes to renamed as 'OLD:NEW'",
    )
    parser.add_argument(
        "--transform",
        metavar="COLUMNS-MAP",
        type=parse_column_transforms,
        help="comma-separated list of column names/indexes to transform as 'NEW:FUNC=OLD'",
    )
    parser.add_argument(
        "--delete",
        metavar="COLUMNS",
        type=parse_columns,
        help="comma-separated list of column names/indexes to delete",
    )
    parser.add_argument(
        "-o",
        "--output",
        nargs="*",
        metavar="FILENAME",
        help="output filename (default: use source file with .numbers)",
    )
    parser.add_argument("--password", help="Password for encrypted documents")
    parser.add_argument("csvfile", nargs="*", help="CSV file to convert")
    return parser


def main(args_list: list[str] | None = None) -> None:
    """Convert the document and exit."""
    parser = command_line_parser()
    args = parser.parse_args(args_list)

    if args.version:
        print(_get_version())
        sys.exit(0)
    elif len(args.csvfile) == 0:
        print("At least one CSV file is required", file=sys.stderr)
        parser.print_help(sys.stderr)
        sys.exit(1)

    if args.output is None:
        output_filenames = [Path(x).with_suffix(".numbers") for x in args.csvfile]
    else:
        output_filenames = args.output

    if len(args.csvfile) != len(output_filenames):
        print("The number of input and output file names do not match", file=sys.stderr)
        sys.exit(1)

    try:
        for input_filename, output_filename in zip(args.csvfile, output_filenames, strict=False):
            converter = Converter(
                day_first=args.day_first,
                no_header=args.no_header,
                whitespace=args.whitespace,
                reverse=args.reverse,
                date_columns=args.date,
                input_filename=input_filename,
                output_filename=str(output_filename),
                encoding=args.encoding,
                password=args.password,
            )

            converter.transform_columns(args.transform)
            converter.rename_columns(args.rename)
            converter.delete_columns(args.delete)
            converter.save()
    except RuntimeError as e:
        print(e, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()
