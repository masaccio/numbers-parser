:hidetoc: 1

Command-line scripts
####################

When installed from `PyPI <https://pypi.org/project/numbers-parser/>`__,
a command-like script ``cat-numbers`` is installed in Python’s scripts
folder. This script dumps Numbers spreadsheets into Excel-compatible CSV
format, iterating through all the spreadsheets passed on the
command-line.

.. code:: text

   usage: cat-numbers [-h] [-T | -S | -b] [-V] [--debug] [--formulas]
                      [--formatting] [-s SHEET] [-t TABLE] [document ...]

   Export data from Apple Numbers spreadsheet tables

   positional arguments:
     document                 Document(s) to export

   optional arguments:
     -h, --help               show this help message and exit
     -T, --list-tables        List the names of tables and exit
     -S, --list-sheets        List the names of sheets and exit
     -b, --brief              Don't prefix data rows with name of sheet/table (default: false)
     -V, --version
     --debug                  Enable debug output
     --formulas               Dump formulas instead of formula results
     --formatting             Dump formatted cells (durations) as they appear in Numbers
     -s SHEET, --sheet SHEET  Names of sheet(s) to include in export
     -t TABLE, --table TABLE  Names of table(s) to include in export

Note: ``--formatting`` will return different capitalization for 12-hour
times due to differences between Numbers’ representation of these dates
and ``datetime.strftime``. Numbers in English locales displays 12-hour
times with ‘am’ and ‘pm’, but ``datetime.strftime`` on macOS at least
cannot return lower-case versions of AM/PM.