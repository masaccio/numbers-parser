import argparse

from numbers_parser import Document


def command_arguments():
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
    parser.add_argument(
        "-s", "--sheet", nargs="*", help="Names of sheet(s) to include in export"
    )
    parser.add_argument(
        "-t", "--table", nargs="*", help="Names of table(s) to include in export"
    )
    parser.add_argument("document", nargs="*", help="Document(s) to export")
    return parser.parse_args()


def print_sheet_names(filenames):
    for filename in filenames:
        for sheet in Document(filename).sheets():
            print(f"{filename}: {sheet.name}")


def print_table_names(filenames):
    for filename in filenames:
        for sheet in Document(filename).sheets():
            for table in sheet.tables():
                print(f"{filename}: {sheet.name}: {table.name}")


args = command_arguments()
if args.list_sheets:
    print_sheet_names(args.document)
elif args.list_tables:
    print_table_names(args.document)
else:
    for filename in args.document:
        for sheet in Document(filename).sheets():
            for table in sheet.tables():
                for row in table.data:
                    print(
                        f"{filename}: {sheet.name}: {table.name}:",
                        ",".join([str(s) if s is not None else "" for s in row]),
                    )
