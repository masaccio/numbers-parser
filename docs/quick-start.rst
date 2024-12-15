Quick Start
-----------

Reading documents:

.. code:: python

   >>> from numbers_parser import Document
   >>> doc = Document("mydoc.numbers")
   >>> sheets = doc.sheets
   >>> tables = sheets[0].tables
   >>> rows = tables[0].rows()

Sheets and tables are iterables that can be indexed using either an
integer index or using the name of the sheet/table:

.. code:: python

   >>> doc.sheets[0].name
   'Sheet 1'
   >>> doc.sheets["Sheet 1"].name
   'Sheet 1'
   >>> doc.sheets[0].tables[0].name
   'Table 1'
   >>> doc.sheets[0].tables["Table 1"].name
   'Table 1'

``Table`` objects have a ``rows`` method which contains a nested list
with an entry for each row of the table. Each row is itself a list of
the column values.

.. code:: python


   >>> data = sheets["Sheet 1"].tables["Table 1"].rows()
   >>> data[0][0]
   <numbers_parser.cell.EmptyCell object at 0x1022b5710>
   >>> data[1][0]
   <numbers_parser.cell.TextCell object at 0x101eb6790>
   >>> data[1][0].value
   'Debit'

Cell Data
^^^^^^^^^

Cells are objects with a common base class of ``Cell``. All cell types
have a property ``value`` which returns the contents of the cell as a
python datatype. Available cell types are:

+---------------+------------------------+---------------------------------+
| Cell type     | value type             | Additional properties           |
+===============+========================+=================================+
| NumberCell    | ``float``              |                                 |
+---------------+------------------------+---------------------------------+
| TextCell      | ``str``                |                                 |
+---------------+------------------------+---------------------------------+
| RichTextCell  | ``str``                | See `Rich text                  |
|               |                        | <https://masaccio.github.io/    |
|               |                        | numbers-parser/api/cells.html#  |
|               |                        | numbers_parser.RichTextCell>`__ |
+---------------+------------------------+---------------------------------+
| EmptyCell     | ``None``               |                                 |
+---------------+------------------------+---------------------------------+
| BoolCell      | ``bool``               |                                 |
+---------------+------------------------+---------------------------------+
| DateCell      | ``datetime.datetime``  |                                 |
+---------------+------------------------+---------------------------------+
| DurationCell  | ``datetime.timedelta`` |                                 |
+---------------+------------------------+---------------------------------+
| ErrorCell     | ``None``               |                                 |
+---------------+------------------------+---------------------------------+
| MergedCell    | ``None``               | See `Merged cells               |
|               |                        | <https://masaccio.github.io/    |
|               |                        | numbers-parser/api/cells.html   |
|               |                        | #numbers_parser.MergedCell>`__  |
+---------------+------------------------+---------------------------------+

Cell references can be either zero-offset row/column integers or an
Excel/Numbers A1 notation. Where cell values are not ``None`` the
property ``formatted_value`` returns the cell value as a ``str`` as
displayed in Numbers. Cells that have no values in a table are
represented as ``EmptyCell`` and cells containing evaluation errors of
any kind ``ErrorCell``.

.. code:: python

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

Pandas Support
^^^^^^^^^^^^^^

Since the return value of ``rows()`` is a list of lists, you can pass
this directly to pandas. Assuming you have a Numbers table with a single
header which contains the names of the pandas series you want to create
you can construct a pandas dataframe using:

.. code:: python

   import pandas as pd

   doc = Document("simple.numbers")
   sheets = doc.sheets
   tables = sheets[0].tables
   data = tables[0].rows(values_only=True)
   df = pd.DataFrame(data[1:], columns=data[0])

Writing Numbers Documents
^^^^^^^^^^^^^^^^^^^^^^^^^

Whilst support for writing numbers files has been stable since version
3.4.0, you are highly recommended not to overwrite working Numbers files
and instead save data to a new file.

Cell values are written using
:pages:`Table.write() <api/table.html#numbers_parser.Table.write>` and
``numbers-parser`` will automatically create empty rows and columns
for any cell references that are out of range of the current table.

.. code:: python

   doc = Document("write.numbers")
   sheets = doc.sheets
   tables = sheets[0].tables
   table = tables[0]
   table.write(1, 1, "This is new text")
   table.write("B7", datetime(2020, 12, 25))
   doc.save("new-sheet.numbers")

Additional tables and worksheets can be added to a ``Document`` before saving using 
:pages:`Document.add_sheet() <api/document.html#numbers_parser.Document.add_sheet>` and
:pages:`Sheet.add_table() <api/sheet.html#numbers_parser.Sheet.add_table>` respectively:

.. code:: python

   doc = Document()
   doc.add_sheet("New Sheet", "New Table")
   sheet = doc.sheets["New Sheet"]
   table = sheet.tables["New Table"]
   table.write(1, 1, 1000)
   table.write(1, 2, 2000)
   table.write(1, 3, 3000)
   doc.save("sheet.numbers")


Styles
^^^^^^

.. include:: styles.rst

Cell Data Formatting
^^^^^^^^^^^^^^^^^^^^

Numbers has two different cell formatting types: data formats and custom
formats.

Data formats are presented in Numbers in the Cell tab of the Format pane
and are applied to individual cells. Like Numbers, ``numbers-parsers``
caches formatting information that is identical across multiple cells.
You do not need to take any action for this to happen; this is handled
internally by the package. Changing a data format for cell has no impact
on any other cells.

Cell formats are changed using
:pages:`Table.set_cell_formatting() <api/table.html#numbers_parser.Table.set_cell_formatting>`:

.. code:: python

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

Custom formats are shared across a Document and can be applied to
multiple cells in multiple tables. Editing a custom format changes the
appearance of data in all cells that share that format. You must first
add a custom format to the document using
:pages:`Document.add_custom_format() <api/document.html#numbers_parser.Document.add_custom_format>`
before assigning it to cells using
:pages:`Table.set_cell_formatting() <api/table.html#numbers_parser.Table.set_cell_formatting>`:

.. code:: python

   long_date = doc.add_custom_format(
      name="Long Date", 
      type="datetime", 
      date_time_format="EEEE, d MMMM yyyy"
   )
   table.set_cell_formatting("C1", "custom", format=long_date)

A limited number of currencies are formatted using symbolic notation
rather than an ISO code. These are defined in
``numbers_parser.currencies`` and match the ones chosen by Numbers. For
example, US dollars are referred to as ``US$`` whereas Euros and British
Pounds are referred to using their symbols of ``€`` and ``£``
respectively.

Borders
^^^^^^^

``numbers-parser`` supports reading and writing cell borders, though the
interface for each differs. Individual cells can have each of their four
borders tested, but when drawing new borders, these are set for the
table to allow for drawing borders across multiple cells. Setting the
border of merged cells is not possible unless the edge of the cells is
at the end of the merged region.

Borders are represented using the :pages:`Border <api/border.html>` class
that can be initialized with line width, color and line style. The
current state of a cell border is read using the
:pages:`Cell.border <api/cells.html#numbers_parser.Cell.border>` property
and :pages:`Table.set_cell_border() <api/table.html#numbers_parser.Table.set_cell_border>`
sets the border for a cell edge or a range of cells.
