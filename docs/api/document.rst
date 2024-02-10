:hidetoc: 1

Document Class
##############

.. currentmodule:: numbers_parser

:class:`Document` represents a single Numbers document which contains multiple
sheets. Sheets can contain multiple tables. The contents of the document are read
once on construction of a :class:`Document`:

.. code-block:: python

   >>> from numbers_parser import Document
   >>> doc = Document("mydoc.numbers")
   >>> doc.sheets[0].name
   'Sheet 1'
   >>> table = doc.sheets[0].tables[0]
   >>> table.name
   'Table 1'

A new document is created when :class:`Document` is constructed without a filename argument:

.. code-block:: python

   doc = Document()
   doc.add_sheet("New Sheet", "New Table")
   sheet = doc.sheets["New Sheet"]
   table = sheet.tables["New Table"]
   table.write(1, 1, 1000)
   table.write(1, 2, 2000)
   table.write(1, 3, 3000)
   doc.save("mydoc.numbers")

.. autoclass:: Document
   :members:

