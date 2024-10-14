"""Command-line utility to convert CSV files to Apple Numbers spreadsheets."""

from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from sys import exit, stderr
from typing import NamedTuple, Tuple  # noqa: F401

from dateutil.parser import parse

from numbers_parser import Document, NumbersError, _get_version


class ColumnTransform(NamedTuple):
    """Class for holding a column transformation rule."""

    source: list[str]
    dest: str
    func: callable


@dataclass
class Converter:
    input_filename: str = None
    output_filename: str = None
    date_columns: list = None
    day_first: bool = False
    no_header: bool = False
    reverse: bool = False
    whitespace: bool = None
    encoding: str = "utf-8"

    def __post_init__(self: Converter) -> None:
        """Parse CSV file with Pandas and return a dataframe."""
        self._read_csv()
        self._transform_data()

    def _read_csv(self) -> None:
        """Read a CSV file into the data array."""
        try:
            dialect = csv.excel
            dialect.strict = True
            lineno = 1
            with open(self.input_filename, encoding=self.encoding) as csvfile:
                csvreader = csv.reader(csvfile, dialect=dialect)
                if self.no_header:
                    self.header = None
                else:
                    self.header = next(csvreader)
                    lineno += 1
                self.data = []
                for row in csvreader:
                    self.data += [row]
                    lineno += 1
        except FileNotFoundError as e:
            msg = f"{self.input_filename}: file not found"
            raise RuntimeError(msg) from e
        except csv.Error as e:
            msg = f"{self.input_filename}@{lineno}: {e.args[0]}"
            raise RuntimeError(msg) from e

    def _parse_date(self, x) -> datetime:
        """Parse a date string and return a datetime."""
        return parse(x, dayfirst=self.day_first).replace(tzinfo=timezone.utc)

    def _transform_data(self):
        """Apply type transformations to the data based in current configuration."""
        # Convert data rows to dicts. csv.DictReader is not enough as we support CSV
        # files with no header.
        if self.no_header:
            self.header = [x for x in range(len(self.data[0]))]
        self.data = [{k: v for k, v in dict(zip(self.header, row)).items()} for row in self.data]

        if self.reverse:
            self.data = list(reversed(self.data))
        if self.date_columns is not None:
            is_date_column = {x: True if x in self.date_columns else False for x in self.header}
        for row in self.data:
            for k, v in row.items():
                if self.whitespace:
                    row[k] = re.sub(r"\s+", " ", v.strip())
                if self.date_columns is not None and is_date_column[k]:
                    row[k] = self._parse_date(v)
                else:
                    # Attempt to coerce value into float
                    try:
                        row[k] = float(v.replace(",", ""))
                    except ValueError:
                        pass

    def rename_columns(self: Converter, mapper: dict) -> None:
        """Rename columns using column map."""
        if mapper is None:
            return
        self.no_header = False
        self.header = [mapper[x] if x in mapper else x for x in self.header]

    def delete_columns(self: Converter, columns: list) -> None:
        """Delete columns from the data."""
        if columns is None:
            return

        missing_columns = list(set(columns) - set(self.header))
        if len(missing_columns) > 0:
            msg = "'" + "', '".join([str(x) for x in missing_columns]) + "'"
            msg += ": cannot delete: column(s) do not exist in CSV"
            raise RuntimeError(msg) from None

        for col in columns:
            del self.header[self.header.index(col)]
        for row in self.data:
            for col in columns:
                del row[col]

    def transform_columns(self: Converter, columns: list[ColumnTransform]) -> None:
        """Perform column transformationstransformations."""
        if columns is None:
            return
        for transform in columns:
            transform.transform(self.header, self.data)
            if transform.dest not in self.header:
                self.header.append(transform.dest)

    def save(self: Converter) -> None:
        """Write dataframe transctions to a Numbers file."""
        doc = Document(num_rows=2, num_cols=2)
        table = doc.sheets[0].tables[0]

        if self.no_header:
            data = []
        else:
            data = [self.header]
        data += [row.values() for row in self.data]

        for row_num, row in enumerate(data):
            for col_num, value in enumerate(row):
                table.write(row_num, col_num, value)
                if isinstance(value, datetime):
                    table.set_cell_formatting(
                        row_num, col_num, "datetime", date_time_format="d MMM yyyy",
                    )

        doc.save(self.output_filename)


class Transformer:
    """Base class for column transformations."""

    def __init__(self: Transformer, source: str, dest: str) -> None:
        self.dest = int(dest) if dest.isnumeric() else dest
        self.sources = [int(x) if x.isnumeric() else x for x in source.split(";")]

    def transform_row(self: Transformer, row: list[str]) -> list[str]:
        """Abstract base method for transforming rows using df.apply()."""
        raise NotImplementedError

    def transform(self: Transformer, header: list[str], data: list[list[str]]) -> list[list[str]]:
        """Column transform to merge columns."""
        if not all(x in header for x in self.sources):
            missing = list(set(self.sources) - set(header))
            msg = "'" + "', '".join([str(x) for x in missing]) + "'"
            msg += ": transform failed: column(s) do not exist in CSV"
            raise RuntimeError(msg)
        for row in data:
            self.transform_row(row)


class MergeTransformer(Transformer):
    """Transformer for column MERGE operations."""

    def transform_row(self: MergeTransformer, row: list[str]) -> list[str]:
        """Merge data in a single row."""
        value = ""
        for col in self.sources:
            if row[col] and not value:
                value = row[col]
        row[self.dest] = value


class NegTransformer(Transformer):
    """Transformer for column NEG operations."""

    def transform_row(self: NegTransformer, row: list[str]) -> list[str]:
        """Select negative values for a row."""
        value = ""
        for col in self.sources:
            if row[col] and not value and float(row[col]) < 0:
                value = abs(float(row[col]))
        row[self.dest] = value


class PosTransformer(Transformer):
    """Transformer for column POS operations."""

    def transform_row(self: PosTransformer, row: list[str]) -> list[str]:
        """Select positive values for a row."""
        value = ""
        for col in self.sources:
            if row[col] and not value and float(row[col]) > 0:
                value = float(row[col])
        row[self.dest] = value


class LookupTransformer(Transformer):
    """Transformer for column LOOKUP operations."""

    def __init__(self: Transformer, source: str, dest: str) -> None:
        super().__init__(source, dest)

        if len(self.sources) != 2:
            msg = f"'{self.sources}' LOOKUP must have exactly 2 arguments"
            raise RuntimeError(msg) from None

        (source, map_filname) = self.sources
        self.sources = [source]

        if not Path(map_filname).exists():
            msg = f"{map_filname}: no such file or directory"
            raise RuntimeError(msg) from None

        try:
            doc = Document(map_filname)
            table = doc.sheets[0].tables[0]
            self.lookup_map = {
                table.cell(row_num, 0).value: table.cell(row_num, 1).value
                for row_num in range(table.num_rows)
            }
        except NumbersError as e:
            msg = f"{map_filname}: {e!r}"
            raise RuntimeError(msg) from e

    def transform_row(self: LookupTransformer, row: list[str]) -> list[str]:
        """Column transform to map values based on a lookup table."""
        matches = [
            {"value": v, "len": len(k)}
            for k, v in self.lookup_map.items()
            if k.lower() in row[self.sources[0]].lower()
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
            old = int(old) if old.isnumeric() else old
            mapper[old] = new
    except csv.Error as e:
        msg = f"'{arg}': malformed CSV string"
        raise argparse.ArgumentTypeError(msg) from e
    else:
        return mapper


def parse_column_transforms(arg: str) -> list[ColumnTransform]:
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
            transforms.append(TRANSFORMERS[func.lower()](source, dest))
    except csv.Error as e:
        msg = f"'{arg}': malformed CSV string"
        raise argparse.ArgumentTypeError(msg) from e
    else:
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
    parser.add_argument("csvfile", nargs="*", help="CSV file to convert")
    return parser


def main() -> None:
    """Convert the document and exit."""
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(_get_version())
        exit(0)
    elif len(args.csvfile) == 0:
        print("At least one CSV file is required", file=stderr)
        parser.print_help(stderr)
        exit(1)

    if args.output is None:
        output_filenames = [Path(x).with_suffix(".numbers") for x in args.csvfile]
    else:
        output_filenames = args.output

    if len(args.csvfile) != len(output_filenames):
        print("The numbers of input and output file names do not match", file=stderr)
        exit(1)

    try:
        for input_filename, output_filename in zip(args.csvfile, output_filenames):
            converter = Converter(
                day_first=args.day_first,
                no_header=args.no_header,
                whitespace=args.whitespace,
                reverse=args.reverse,
                date_columns=args.date,
                input_filename=input_filename,
                output_filename=output_filename,
                encoding=args.encoding,
            )

            converter.transform_columns(args.transform)
            converter.rename_columns(args.rename)
            converter.delete_columns(args.delete)
            converter.save()
    except RuntimeError as e:
        print(e, file=stderr)
        exit(1)


if __name__ == "__main__":  # pragma: no cover
    # execute only if run as a script
    main()
