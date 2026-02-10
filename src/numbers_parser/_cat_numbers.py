import argparse
import csv
import logging
import sys

from sigfig import round as sigfig

from numbers_parser import (
    Document,
    ErrorCell,
    FileError,
    FileFormatError,
    NumberCell,
    UnsupportedError,
    _get_version,
)
from numbers_parser import __name__ as numbers_parser_name
from numbers_parser.constants import MAX_SIGNIFICANT_DIGITS
from numbers_parser.experimental import ExperimentalFeatures, enable_experimental_feature

logger = logging.getLogger(numbers_parser_name)


def experimental_feature_choice(s: str) -> ExperimentalFeatures:
    try:
        return ExperimentalFeatures[s]
    except KeyError:
        msg = f"invalid experimental feature: {s}"
        raise argparse.ArgumentTypeError(msg)  # noqa: B904


def command_line_parser():
    parser = argparse.ArgumentParser(
        description="Export data from Apple Numbers spreadsheet tables",
    )
    commands = parser.add_mutually_exclusive_group()
    commands.add_argument(
        "-T",
        "--list-tables",
        action="store_true",
        help="List the names of tables and exit",
    )
    commands.add_argument(
        "-S",
        "--list-sheets",
        action="store_true",
        help="List the names of sheets and exit",
    )
    commands.add_argument(
        "-b",
        "--brief",
        action="store_true",
        default=False,
        help="Don't prefix data rows with name of sheet/table (default: false)",
    )
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument(
        "--formulas",
        action="store_true",
        help="Dump formulas instead of formula results",
    )
    parser.add_argument(
        "--formatting",
        action="store_true",
        help="Dump formatted cells (durations) as they appear in Numbers",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        action="append",
        help="Names of sheet(s) to include in export",
    )
    parser.add_argument(
        "-t",
        "--table",
        action="append",
        help="Names of table(s) to include in export",
    )
    parser.add_argument("document", nargs="*", help="Document(s) to export")
    parser.add_argument("--debug", default=False, action="store_true", help="Enable debug logging")
    experimental_choices = [
        f.name for f in ExperimentalFeatures if f is not ExperimentalFeatures.NONE
    ]
    parser.add_argument(
        "--experimental",
        type=experimental_feature_choice,
        choices=[ExperimentalFeatures[name] for name in experimental_choices],
        help=argparse.SUPPRESS,
    )
    return parser


def print_sheet_names(filename) -> None:
    for sheet in Document(filename).sheets:
        print(f"{filename}: {sheet.name}")


def print_table_names(filename) -> None:
    for sheet in Document(filename).sheets:
        for table in sheet.tables:
            print(f"{filename}: {sheet.name}: {table.name}")


def cell_as_string(args, cell):
    if isinstance(cell, ErrorCell) and not (args.formulas):
        return "#REF!"
    if args.formulas and cell.formula is not None:
        return cell.formula
    if args.formatting and cell.formatted_value is not None:
        return cell.formatted_value
    if isinstance(cell, NumberCell):
        return sigfig(cell.value, sigfigs=MAX_SIGNIFICANT_DIGITS, warn=False)
    if cell.value is None:
        return ""
    return str(cell.value)


def print_table(args, filename) -> None:
    writer = csv.writer(sys.stdout, dialect="excel")
    for sheet in Document(filename).sheets:
        if args.sheet is not None and sheet.name not in args.sheet:
            continue
        for table in sheet.tables:
            if args.table is not None and table.name not in args.table:
                continue
            for row in table.rows():
                cells = [cell_as_string(args, cell) for cell in row]
                if not args.brief:
                    sys.stdout.write(f"{filename}: {sheet.name}: {table.name}: ")
                writer.writerow(cells)


def main() -> None:
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(_get_version())
    elif len(args.document) == 0:
        parser.print_help()
    else:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
        logger.addHandler(handler)
        if args.debug:
            logger.setLevel("DEBUG")
        else:
            logger.setLevel("ERROR")
        if args.experimental:
            enable_experimental_feature(args.experimental)
        for filename in args.document:
            try:
                if args.list_sheets:
                    print_sheet_names(filename)
                elif args.list_tables:
                    print_table_names(filename)
                else:
                    print_table(args, filename)
            except (FileFormatError, FileError, UnsupportedError) as e:  # noqa: PERF203
                err_str = str(e)
                if filename in err_str:
                    print(err_str, file=sys.stderr)
                else:
                    print(f"{filename}: {err_str}", file=sys.stderr)
                sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()
