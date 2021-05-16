# vi: ft=python

import argparse

from numbers_parser import Document, _get_version


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
    parser.add_argument("-s", "--sheet", action="append", help="Names of sheet(s) to include in export")
    parser.add_argument("-t", "--table", action="append", help="Names of table(s) to include in export")
    parser.add_argument("document", nargs="*", help="Document(s) to export")
    return parser


def print_sheet_names(filename):
    for sheet in Document(filename).sheets():
        print(f"{filename}: {sheet.name}")


def print_table_names(filename):
    for sheet in Document(filename).sheets():
        for table in sheet.tables():
            print(f"{filename}: {sheet.name}: {table.name}")


def print_table(args, filename):
    for sheet in Document(filename).sheets():
        if args.sheet is not None and sheet.name not in args.sheet:
            continue
        for table in sheet.tables():
            if args.table is not None and table.name not in args.table:
                continue
            for row in table.data:
                cols = ",".join([str(s) if s is not None else "" for s in row])
                if args.brief:
                    print(cols)
                else:
                    print(f"{filename}: {sheet.name}: {table.name}:", cols)

def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(_get_version())
    elif len(args.document) == 0:
        parser.print_help()
    else:
        for filename in args.document:
            if args.list_sheets:
                print_sheet_names(filename)
            elif args.list_tables:
                print_table_names(filename)
            else:
                print_table(args, filename)

if __name__ == "__main__":
    # execute only if run as a script
    main()
