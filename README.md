# numbers-parser

[![Test Status](https://github.com/masaccio/numbers-parser/actions/workflows/run-all-tests.yml/badge.svg)](https://github.com/masaccio/numbers-parser/actions/workflows/run-all-tests.yml)[![Security Checks](https://github.com/masaccio/numbers-parser/actions/workflows/codeql.yml/badge.svg)](https://github.com/masaccio/numbers-parser/actions/workflows/codeql.yml)[![Code Coverage](https://codecov.io/gh/masaccio/numbers-parser/branch/main/graph/badge.svg?token=EKIUFGT05E)](https://codecov.io/gh/masaccio/numbers-parser)[![PyPI Version](https://badge.fury.io/py/numbers-parser.svg)](https://badge.fury.io/py/numbers-parser)

`numbers-parser` is a Python module for parsing [Apple Numbers](https://www.apple.com/numbers/)`.numbers` files. It supports Numbers files
generated by Numbers version 10.3, and up with the latest tested version being 14.1
(current as of June 2024).

It supports and is tested against Python versions from 3.9 onwards. It is not compatible
with earlier versions of Python.

## Installation

A pre-requisite for this package is [python-snappy](https://pypi.org/project/python-snappy/) which will be installed by Python automatically, but python-snappy also requires binary libraries for snappy compression.

The most straightforward way to install the binary dependencies is to use
[Homebrew](https://brew.sh) and source Python from Homebrew rather than from macOS as described
in the [python-snappy github](https://github.com/andrix/python-snappy). Using [pipx](https://pipx.pypa.io/stable/installation/) for package management is also strongly recommended:

```bash
brew install snappy python3 pipx
pipx install numbers-parser
```

For Linux (your package manager may be different):

```bash
sudo apt-get -y install libsnappy-dev
```

On Windows, you will need to either arrange for snappy to be found for VSC++ or you can install python
[pre-compiled binary libraries](https://github.com/cgohlke/win_arm64-wheels/) which are only available
for Windows on Arm. There appear to be no x86 pre-compiled packages for Windows.

```text
pip install python_snappy-0.6.1-cp312-cp312-win_arm64.whl
```

## Quick Start

Reading documents:

```python
>>> from numbers_parser import Document
>>> doc = Document("mydoc.numbers")
>>> sheets = doc.sheets
>>> tables = sheets[0].tables
>>> rows = tables[0].rows()
```

Sheets and tables are iterables that can be indexed using either an
integer index or using the name of the sheet/table:

```python
>>> doc.sheets[0].name
'Sheet 1'
>>> doc.sheets["Sheet 1"].name
'Sheet 1'
>>> doc.sheets[0].tables[0].name
'Table 1'
>>> doc.sheets[0].tables["Table 1"].name
'Table 1'
```

`Table` objects have a `rows` method which contains a nested list
with an entry for each row of the table. Each row is itself a list of
the column values.

```python
>>> data = sheets["Sheet 1"].tables["Table 1"].rows()
>>> data[0][0]
<numbers_parser.cell.EmptyCell object at 0x1022b5710>
>>> data[1][0]
<numbers_parser.cell.TextCell object at 0x101eb6790>
>>> data[1][0].value
'Debit'
```

### Cell Data

Cells are objects with a common base class of `Cell`. All cell types
have a property `value` which returns the contents of the cell as a
python datatype. Available cell types are:

| Cell type    | value type           | Additional properties                                                                                  |
|--------------|----------------------|--------------------------------------------------------------------------------------------------------|
| NumberCell   | `float`              |                                                                                                        |
| TextCell     | `str`                |                                                                                                        |
| RichTextCell | `str`                | See [Rich text](https://masaccio.github.io/numbers-parser/api/cells.html#numbers_parser.RichTextCell)  |
| EmptyCell    | `None`               |                                                                                                        |
| BoolCell     | `bool`               |                                                                                                        |
| DateCell     | `datetime.datetime`  |                                                                                                        |
| DurationCell | `datetime.timedelta` |                                                                                                        |
| ErrorCell    | `None`               |                                                                                                        |
| MergedCell   | `None`               | See [Merged cells](https://masaccio.github.io/numbers-parser/api/cells.html#numbers_parser.MergedCell) |

Cell references can be either zero-offset row/column integers or an
Excel/Numbers A1 notation. Where cell values are not `None` the
property `formatted_value` returns the cell value as a `str` as
displayed in Numbers. Cells that have no values in a table are
represented as `EmptyCell` and cells containing evaluation errors of
any kind `ErrorCell`.

```python
>>> table.cell(1,0)
<numbers_parser.cell.TextCell object at 0x1019ade50>
>>> table.cell(1,0).value
'Debit'
>>> table.cell("B2")
<numbers_parser.cell.NumberCell object at 0x103a99790>
>>> table.cell("B2").value
1234.5
>>> table.cell("B2").formatted_value
'£1,234.50'
```

### Pandas Support

Since the return value of `rows()` is a list of lists, you can pass
this directly to pandas. Assuming you have a Numbers table with a single
header which contains the names of the pandas series you want to create
you can construct a pandas dataframe using:

```python
import pandas as pd

doc = Document("simple.numbers")
sheets = doc.sheets
tables = sheets[0].tables
data = tables[0].rows(values_only=True)
df = pd.DataFrame(data[1:], columns=data[0])
```

### Writing Numbers Documents

Whilst support for writing numbers files has been stable since version
3.4.0, you are highly recommended not to overwrite working Numbers files
and instead save data to a new file.

Cell values are written using
[Table.write()](https://masaccio.github.io/numbers-parser/api/table.html#numbers_parser.Table.write) and
`numbers-parser` will automatically create empty rows and columns
for any cell references that are out of range of the current table.

```python
doc = Document("write.numbers")
sheets = doc.sheets
tables = sheets[0].tables
table = tables[0]
table.write(1, 1, "This is new text")
table.write("B7", datetime(2020, 12, 25))
doc.save("new-sheet.numbers")
```

Additional tables and worksheets can be added to a `Document` before saving using
[Document.add_sheet()](https://masaccio.github.io/numbers-parser/api/document.html#numbers_parser.Document.add_sheet) and
[Sheet.add_table()](https://masaccio.github.io/numbers-parser/api/sheet.html#numbers_parser.Sheet.add_table) respectively:

```python
doc = Document()
doc.add_sheet("New Sheet", "New Table")
sheet = doc.sheets["New Sheet"]
table = sheet.tables["New Table"]
table.write(1, 1, 1000)
table.write(1, 2, 2000)
table.write(1, 3, 3000)
doc.save("sheet.numbers")
```

### Styles

`numbers_parser` currently only supports paragraph styles and cell
styles. The following styles are supported:

- font attributes: bold, italic, underline, strikethrough
- font selection and size
- text foreground color
- horizontal and vertical alignment
- cell background color
- cell background images
- cell indents (first line, left, right, and text inset)

Numbers conflates style attributes that can be stored in paragraph
styles (the style menu in the text panel) with the settings that are
available on the Style tab of the Text panel. Some attributes in Numbers
are not applied to new cells when a style is applied.

To keep the API simple, `numbers-parser` packs all styling into a single
[Style](https://masaccio.github.io/numbers-parser/api/style.html) object. When a document is saved, the attributes
not stored in a paragraph style are applied to each cell that includes it.

Styles are read from cells using the
[Cell.style](https://masaccio.github.io/numbers-parser/api/cells.html#numbers_parser.Cell.style) property and you can
add new styles with
[Document.add_style](https://masaccio.github.io/numbers-parser/api/document.html#numbers_parser.Document.add_style).

```python
red_text = doc.add_style(
    name="Red Text",
    font_name="Lucida Grande",
    font_color=RGB(230, 25, 25),
    font_size=14.0,
    bold=True,
    italic=True,
    alignment=Alignment("right", "top"),
)
table.write("B2", "Red", style=red_text)
table.set_cell_style("C2", red_text)
```

### Cell Data Formatting

Numbers has two different cell formatting types: data formats and custom
formats.

Data formats are presented in Numbers in the Cell tab of the Format pane
and are applied to individual cells. Like Numbers, `numbers-parsers`
caches formatting information that is identical across multiple cells.
You do not need to take any action for this to happen; this is handled
internally by the package. Changing a data format for cell has no impact
on any other cells.

Cell formats are changed using
[Table.set_cell_formatting()](https://masaccio.github.io/numbers-parser/api/table.html#numbers_parser.Table.set_cell_formatting):

```python
table.set_cell_formatting(
   "C1",
   "datetime",
   date_time_format="EEEE, d MMMM yyyy"
)
table.set_cell_formatting(
   0,
   4,
   "number",
   decimal_places=3,
   negative_style=NegativeNumberStyle.RED
)
```

Custom formats are shared across a Document and can be applied to
multiple cells in multiple tables. Editing a custom format changes the
appearance of data in all cells that share that format. You must first
add a custom format to the document using
[Document.add_custom_format()](https://masaccio.github.io/numbers-parser/api/document.html#numbers_parser.Document.add_custom_format)
before assigning it to cells using
[Table.set_cell_formatting()](https://masaccio.github.io/numbers-parser/api/table.html#numbers_parser.Table.set_cell_formatting):

```python
long_date = doc.add_custom_format(
   name="Long Date",
   type="datetime",
   date_time_format="EEEE, d MMMM yyyy"
)
table.set_cell_formatting("C1", "custom", format=long_date)
```

A limited number of currencies are formatted using symbolic notation
rather than an ISO code. These are defined in
`numbers_parser.currencies` and match the ones chosen by Numbers. For
example, US dollars are referred to as `US$` whereas Euros and British
Pounds are referred to using their symbols of `€` and `£`
respectively.

### Borders

`numbers-parser` supports reading and writing cell borders, though the
interface for each differs. Individual cells can have each of their four
borders tested, but when drawing new borders, these are set for the
table to allow for drawing borders across multiple cells. Setting the
border of merged cells is not possible unless the edge of the cells is
at the end of the merged region.

Borders are represented using the [Border](https://masaccio.github.io/numbers-parser/api/border.html) class
that can be initialized with line width, color and line style. The
current state of a cell border is read using the
[Cell.border](https://masaccio.github.io/numbers-parser/api/cells.html#numbers_parser.Cell.border) property
and [Table.set_cell_border()](https://masaccio.github.io/numbers-parser/api/table.html#numbers_parser.Table.set_cell_border)
sets the border for a cell edge or a range of cells.

## API

For more examples and details of all available classes and methods,
see the [full API docs](https://masaccio.github.io/numbers-parser/).

## Command-line scripts

When installed from [PyPI](https://pypi.org/project/numbers-parser/),
a number of command-line scripts are installed:

- `cat-numbers`: converts Numbers documents into CSV
- `csv2numbers`: converts CSV files to Numbers documents
- `unpack-numbers`: converts Numbers documents into JSON files for debug purposes

### cat-numbers

This script dumps Numbers spreadsheets into Excel-compatible CSV
format, iterating through all the spreadsheets passed on the
command-line.

```text
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
```

Note: `--formatting` will return different capitalization for 12-hour
times due to differences between Numbers’ representation of these dates
and `datetime.strftime`. Numbers in English locales displays 12-hour
times with ‘am’ and ‘pm’, but `datetime.strftime` on macOS at least
cannot return lower-case versions of AM/PM.

### csv2numbers

This script converts Excel-compatible CSV files into Numbers documents. Output files
can optionally be provided, but is none are provided, the output is created by replacing
the input’s files suffix with .numbers. For example:

```text
csv2numbers file1.csv file2.csv -o file1.numbers file2.numbers
```

Columns of data can have a number of transformations applied to them. The primary use-
case intended for `csv2numbers` is converting banking exports to well-formatted
spreadsheets.

```text
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
```

The following options affecting the output of the entire file. The default for each is always false.

- `--whitespace`: strip whitespace from beginning and end of strings and collapse other whitespace into single space
- `--reverse`: reverse the order of the data rows
- `--no-header`: CSV file has no header row
- ``--day-first`: dates are represented day first in the CSV file

`csv2numbers` can also perform column manipulation. Columns can be identified using their name if the CSV file has a header or using a column index. Columns are zero-indexed and names and indices can be used together on the same command-line. When multiple columns are required, you can specify them using comma-separated values. The format for these arguments, like for the CSV file itself, the Excel dialect.

#### Deleting columns

Delete columns using `--delete`. The names or indices of the columns to delete are specified as comma-separated values:

```text
csv2numbers file1.csv --delete=Account,3
```

#### Renaming columns

Rename columns using `--rename`. The current column name and new column name are separated by a `:` and each renaming is specified as comma-separated values:

```text
csv2numbers file1.csv --rename=2:Account,"Paid In":Amount
```

#### Date columns

The `--date` option identifies a comma-separated list of columns that should be parsed as dates. Use `--day-first` where the day and month is ambiguous anf the day comes first rather than the month.

#### Transforming columns

Columns can be merged and new columns created using simple functions. The –transform option takes a comma-seperated list of transformations of the form NEW:FUNC=OLD. Supported functions are:

| Function   | Arguments                   | Description                                                                                                                                                                                                                                                                                                                                           |
|------------|-----------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| MERGE      | dest=MERGE:source           | The dest column is writen with values from one or more columns<br/>indicated by source. For multiple columns, which are separated<br/>by ;, the first empty value is chosen.                                                                                                                                                                          |
| NEG        | dest=NEG:source             | The dest column contains absolute values of any column that is<br/>negative. This is useful for isolating debits from account<br/>exports.                                                                                                                                                                                                            |
| POS        | dest=NEG:source             | The dest column contains values of any column that is<br/>positive. This is useful for isolating credits from account<br/>exports.                                                                                                                                                                                                                    |
| LOOKUP     | dest=LOOKUP:source;filename | A lookup map is read from filename which must be an Apple<br/>Numbers file containing a single table of two columns. The table<br/>is used to match agsinst source, searching the first column<br/>for matches and writing the corresponding value from the second<br/>column to dest. Values are chosen based on the longest<br/>matching substring. |

Examples:

```text
csv2numbers --transform="Paid In"=POS:Amount,Withdrawn=NEG:Amount file1.csv
csv2numbers --transform='Category=LOOKUP:Transaction;mapping.numbers' file1.csv
```

## Limitations

Current known limitations of `numbers-parser` which may be implemented in the future are:

- Table styles that allow new tables to adopt a style across the whole
  table are not suppported
- Creating cells of type `BulletedTextCell` is not supported
- New tables are inserted with a fixed offset below the last table in a
  worksheet which does not take into account title or caption size
- Captions can be created and edited as of numbers-parser version 4.12, but cannot
  be styled. New captions adopt the first caption style available in the current
  document
- Formulas cannot be written to a document
- Pivot tables are unsupported and saving a document with a pivot table issues
  a UnsupportedWarning (see [issue 73](https://github.com/masaccio/numbers-parser/issues/73) for details).

The following limitations are expected to always remain:

- New sheets insert tables with formats copied from the first table in
  the previous sheet rather than default table formats
- Due to a limitation in Python’s
  [ZipFile](https://docs.python.org/3/library/zipfile.html), Python
  versions older than 3.11 do not support image filenames with UTF-8 characters
  [Cell.add_style.bg_image()](https://masaccio.github.io/numbers-parser/api/sheet.html#numbers_parser.Style) returns
  `None` for such files and issues a `RuntimeWarning`
  (see [issue 69](https://github.com/masaccio/numbers-parser/issues/69) for details).
- Password-encrypted documents cannot be opened. You must first re-save without
  a password to read (see [issue 88](https://github.com/masaccio/numbers-parser/issues/88) for details).
  A UnsupportedError exception is raised when such documents are opened.

## License

All code in this repository is licensed under the [MIT License](https://github.com/masaccio/numbers-parser/blob/master/LICENSE.rst).
