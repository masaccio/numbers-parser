Command-line scripts
--------------------

When installed from `PyPI <https://pypi.org/project/numbers-parser/>`__,
a number of command-line scripts are installed:

- ``cat-numbers``: converts Numbers documents into CSV
- ``csv2numbers``: converts CSV files to Numbers documents
- ``unpack-numbers``: converts Numbers documents into JSON files for debug purposes


cat-numbers
^^^^^^^^^^^

This script dumps Numbers spreadsheets into Excel-compatible CSV format, iterating through all the spreadsheets passed on the command-line.

.. code:: text

    usage: cat-numbers [-h] [-T | -S | -b] [-V] [--formulas] [--formatting]
                      [-s SHEET] [-t TABLE] [--debug]
                      [document ...]

    Export data from Apple Numbers spreadsheet tables

    positional arguments:
      document              Document(s) to export

    options:
      -h, --help            show this help message and exit
      -T, --list-tables     List the names of tables and exit
      -S, --list-sheets     List the names of sheets and exit
      -b, --brief           Don't prefix data rows with name of sheet/table
                            (default: false)
      -V, --version
      --formulas            Dump formulas instead of formula results
      --formatting          Dump formatted cells (durations) as they appear
                            in Numbers
      -s SHEET, --sheet SHEET
                            Names of sheet(s) to include in export
      -t TABLE, --table TABLE
                            Names of table(s) to include in export
      --debug               Enable debug logging

Note: ``--formatting`` will return different capitalization for 12-hour times due to differences between Numbers’ representation of these dates and ``datetime.strftime``. Numbers in English locales displays 12-hour times with ‘am’ and ‘pm’, but ``datetime.strftime`` on macOS at least cannot return lower-case versions of AM/PM.

csv2numbers
^^^^^^^^^^^

This script converts Excel-compatible CSV files into Numbers documents. Output files can optionally be provided, but is none are provided, the output is created by replacing the input's files suffix with `.numbers`. For example:

.. code:: text

  csv2numbers file1.csv file2.csv -o file1.numbers file2.numbers

Columns of data can have a number of transformations applied to them. The primary use- case intended for ``csv2numbers`` is converting banking exports to well-formatted spreadsheets.

.. code:: text

  usage: csv2numbers [-h] [-V] [--whitespace] [--reverse] [--no-header]
                     [--day-first] [--date COLUMNS] [--rename COLUMNS-MAP]
                     [--transform COLUMNS-MAP] [--delete COLUMNS]
                     [-o [FILENAME ...]]
                     [csvfile ...]

  positional arguments:
    csvfile               CSV file to convert

  options:
    -h, --help            show this help message and exit
    -V, --version
    --whitespace          strip whitespace from beginning and end of strings
                          and collapse other whitespace into single space
                          (default: false)
    --reverse             reverse the order of the data rows (default:
                          false)
    --no-header           CSV file has no header row (default: false)
    --day-first           dates are represented day first in the CSV file
                          (default: false)
    --date COLUMNS        comma-separated list of column names/indexes to
                          parse as dates
    --rename COLUMNS-MAP  comma-separated list of column names/indexes to
                          renamed as 'OLD:NEW'
    --transform COLUMNS-MAP
                          comma-separated list of column names/indexes to
                          transform as 'NEW:FUNC=OLD'
    --delete COLUMNS      comma-separated list of column names/indexes to
                          delete
    -o [FILENAME ...], --output [FILENAME ...]
                          output filename (default: use source file with
                          .numbers)

The following options affecting the output of the entire file. The default for each is always false.

- ``--whitespace``: strip whitespace from beginning and end of strings and collapse other whitespace into single space
- ``--reverse``: reverse the order of the data rows
- ``--no-header``: CSV file has no header row
- ```--day-first``: dates are represented day first in the CSV file

``csv2numbers`` can also perform column manipulation. Columns can be identified using their name if the CSV file has a header or using a column index. Columns are zero-indexed and names and indices can be used together on the same command-line. When multiple columns are required, you can specify them using comma-separated values. The format for these arguments, like for the CSV file itself, the Excel dialect.

Deleting columns
""""""""""""""""

Delete columns using ``--delete``. The names or indices of the columns to delete are specified as comma-separated values:

.. code:: text

  csv2numbers file1.csv --delete=Account,3

Renaming columns
"""""""""""""""""

Rename columns using ``--rename``. The current column name and new column name are separated by a ``:`` and each renaming is specified as comma-separated values:

.. code:: text

  csv2numbers file1.csv --rename=2:Account,"Paid In":Amount

Date columns
"""""""""""""

The ``--date`` option identifies a comma-separated list of columns that should be parsed as dates. Use ``--day-first`` where the day and month is ambiguous anf the day comes first rather than the month.

Transforming columns
"""""""""""""""""""""

Columns can be merged and new columns created using simple functions. The `--transform` option takes a comma-seperated list of transformations of the form `NEW:FUNC=OLD`. Supported functions are:

+-------------+-------------------------------+------------------------------------------------------------------+
| Function    | Arguments                     | Description                                                      |
+-------------+-------------------------------+------------------------------------------------------------------+
| `MERGE`     | `dest=MERGE:source`           | The `dest` column is writen with values from one or more columns |
|             |                               | indicated by `source`. For multiple columns, which are separated |
|             |                               | by `;`, the first empty value is chosen.                         |
+-------------+-------------------------------+------------------------------------------------------------------+
| `NEG`       | `dest=NEG:source`             | The `dest` column contains absolute values of any column that is |
|             |                               | negative. This is useful for isolating debits from account       |
|             |                               | exports.                                                         |
+-------------+-------------------------------+------------------------------------------------------------------+
| `POS`       | `dest=NEG:source`             | The `dest` column contains values of any column that is          |
|             |                               | positive. This is useful for isolating credits from account      |
|             |                               | exports.                                                         |
+-------------+-------------------------------+------------------------------------------------------------------+
| `LOOKUP`    | `dest=LOOKUP:source;filename` | A lookup map is read from `filename` which must be an Apple      |
|             |                               | Numbers file containing a single table of two columns. The table |
|             |                               | is used to match agsinst `source`, searching the first column    |
|             |                               | for matches and writing the corresponding value from the second  |
|             |                               | column to `dest`. Values are chosen based on the longest         |
|             |                               | matching substring.                                              |
+-------------+-------------------------------+------------------------------------------------------------------+

Examples:

.. code:: text

  csv2numbers --transform="Paid In"=POS:Amount,Withdrawn=NEG:Amount file1.csv
  csv2numbers --transform='Category=LOOKUP:Transaction;mapping.numbers' file1.csv