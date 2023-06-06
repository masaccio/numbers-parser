import argparse
import csv
import logging
import sys

from numbers_parser import Document, _get_version
from numbers_parser import __name__ as numbers_parser_name
from numbers_parser.exceptions import FileFormatError
from numbers_parser.cell import ErrorCell

logger = logging.getLogger(numbers_parser_name)


def command_line_parser():
    parser = argparse.ArgumentParser(
        description="Export data from Apple Numbers spreadsheet tables"
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
        "-s", "--sheet", action="append", help="Names of sheet(s) to include in export"
    )
    parser.add_argument(
        "-t", "--table", action="append", help="Names of table(s) to include in export"
    )
    parser.add_argument("document", nargs="*", help="Document(s) to export")
    parser.add_argument(
        "--debug", default=False, action="store_true", help="Enable debug logging"
    )
    return parser


def print_sheet_names(filename):
    for sheet in Document(filename).sheets:
        print(f"{filename}: {sheet.name}")


def print_table_names(filename):
    for sheet in Document(filename).sheets:
        for table in sheet.tables:
            print(f"{filename}: {sheet.name}: {table.name}")


def cell_as_string(args, cell):
    if type(cell) == ErrorCell and not (args.formulas):
        return "#REF!"
    elif args.formulas and cell.formula is not None:
        return cell.formula
    elif args.formatting and cell.formatted_value is not None:
        return cell.formatted_value
    elif cell.value is None:
        return ""
    else:
        return str(cell.value)


def print_table(args, filename):
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


def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(_get_version())
    elif len(args.document) == 0:
        parser.print_help()
    else:
        hdlr = logging.StreamHandler()
        hdlr.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
        logger.addHandler(hdlr)
        if args.debug:
            logger.setLevel("DEBUG")
        else:
            logger.setLevel("ERROR")
        for filename in args.document:
            try:
                if args.list_sheets:
                    print_sheet_names(filename)
                elif args.list_tables:
                    print_table_names(filename)
                else:
                    print_table(args, filename)
            except FileFormatError as e:  # pragma: no cover”
                print(f"{filename}:", str(e), file=sys.stderr)
                sys.exit(1)


if __name__ == "__main__":
    # execute only if run as a script
    main()  # pragma: no cover
