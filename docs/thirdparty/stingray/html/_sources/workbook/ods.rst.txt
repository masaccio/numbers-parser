..  #!/usr/bin/env python3

.. _`workbook_ods`:


ODS Workbook
---------------

::

  import logging
  import pprint
  import xml.etree.cElementTree as dom
  import zipfile
  import datetime

  from stingray.workbook.base import Workbook
  import stingray.sheet
  import stingray.cell

..  py:module:: workbook.ods

We should use ``iterparse`` rather than simply parsing the entire XML document.
If the document is large, then we can't hold it all in memory.

..  py:class:: ODS_Workbook

    Extract sheets, rows and cells from a OOO ODS format file.

    In addition to the superclass attributes, some additional unique
    attributes are introduced here.
       
    ..  py:attribute:: zip_archive
   
        A zip archive for this file.
       
    ..  py:attribute:: tables
   
        A mapping that provides sheet names.

::

  class ODS_Workbook( Workbook ):
      """Standard OOO ODS document.
      Locate sheets and rows within a given sheet.
      """
      ODS_NS = {
      "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
      "table":"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
      "text":"urn:oasis:names:tc:opendocument:xmlns:text:1.0",
      }
      date_format = "%Y-%m-%d"
      def __init__( self, name, file_object=None ):
          """Prepare the workbook for reading.
          :param name: File name
          :param file_object: Optional file-like object.  If omitted, the named file is opened.
          """
          super().__init__( name, file_object )
          self.zip_archive= zipfile.ZipFile( file_object or name, "r" )
          self._prepare()

As preparation for reading these files, we locate all the sheet names.

::

      def _prepare( self ):
          self._locate_sheets()

Locating all the sheets is a matter of doing an XPath search for
:samp:`body/spreadsheet/table` and getting the *name* attribute
from the  :samp:`<table name="{name}">` tags.

::

      def _locate_sheets( self ):
          """Create ``tables`` map from name to table."""
          workbook_zip= self.zip_archive.getinfo("content.xml")
          workbook_doc= dom.parse( self.zip_archive.open(workbook_zip) )
          name_attr_id= dom.QName( self.ODS_NS["table"], "name" )
          logging.debug( dom.tostring( workbook_doc.getroot() ) )
          self.tables= dict(
              (t.attrib[name_attr_id],t)
              for t in workbook_doc.findall("office:body/office:spreadsheet/table:table", 
                  namespaces=self.ODS_NS) )

An ``iterparse`` version to locate sheets
would look for start of ``table`` tags and then get
the name attribute from that tag.

..  py:method:: ODS_Workbook.sheets( )

    Return the list of sheets for this workbook.
   
::

      def sheets( self ):
          return self.tables.keys()

We can build an eager :py:class:`sheet.Row` or a :py:class:`sheet.LazyRow` from
the available data.  The eager Row includes the conversions.
The LazyRow defers the conversions to :py:meth:`ODS_Workbook.row_get`.

In ODS documents, the cell's value can be carried in the value attribute or
it can be a mixed content value of the element.  There are three cases.

-   :samp:`<table-cell value-type="{type}" value="{value}">...</table-cell>`

-   :samp:`<table-cell value-type="{type}" date-value="{value}">...</table-cell>`

-   :samp:`<table-cell value-type="{type}">{value}</table-cell>`

..  py:method:: ODS_Workbook.rows_of( sheet )

    Iterate through rows of the given sheet.

::

      def rows_of( self, sheet ):
          """Iterator over rows as a list of Cells for a named worksheet."""
          for r, row_doc in enumerate(
              self.tables[sheet.name].findall( "table:table-row", namespaces=self.ODS_NS ) ):
              row= []
              for c, cell_doc in enumerate( row_doc.findall( "table:table-cell", namespaces=self.ODS_NS ) ):
                  row.append( self.cell(cell_doc) )
              yield row

..  py:method:: ODS_Workbook.row_get( row, attribute )

    Low-level get of a particular attribute from the given row.

::

      def row_get( self, row, attribute ):
          """Create a Cell from the row's data."""
          return row[attribute.position]

Build a subclass of :py:class:`cell.Cell` from the current type name and value.

..  todo:: Refactor this, it feels clunky.

::

      def cell( self, cell_doc ):
          logging.debug( dom.tostring(cell_doc) )
          value_attr_id= dom.QName( self.ODS_NS['office'], 'value' )
          date_attr_id= dom.QName( self.ODS_NS['office'], 'date-value' )
          type_attr_id= dom.QName( self.ODS_NS['office'], 'value-type' )
          # Get the type
          try:
              type_name= cell_doc.attrib[type_attr_id]
          except KeyError:
              return stingray.cell.EmptyCell( '', self )
          value= None
          # Date value as attribute?
          if not value:
              try:
                  value= cell_doc.attrib[date_attr_id]
              except KeyError:
                  pass
          # Other value as attribute?
          if not value:
              try:
                  value= cell_doc.attrib[value_attr_id]
              except KeyError:
                  pass
          # No value attributes, get *all* the text content.
          if not value:
              value= "".join( x for x in cell_doc.itertext() )
          if not value:
              # TODO: Proper warning.
              dom.dump( cell_doc )
          logging.debug( type_name, repr(value) )
          if type_name == "string":
              return stingray.cell.TextCell( value, self )
          elif type_name == "float":
              return stingray.cell.NumberCell( float(value), self )
          elif type_name == "date":
              theDate= datetime.datetime.strptime(
                  value, ODS_Workbook.date_format )
              return stingray.cell.FloatDateCell( theDate, self )
          elif type_name == "boolean":
              return stingray.cell.BooleanCell(
                  float(value.upper()=='TRUE'),  self )
          elif type_name == "empty":
              return stingray.cell.EmptyCell( '', self )
          else: 
              raise Exception( "Unknown cell {0}".format( dom.tostring(cell_doc) ) )
            
An ``iterparse`` version of building a row
would look for start of ``table`` tags and then get
the name attribute from that tag just to locate the right sheet.

Once the sheet was located, then the row and cell tags would be used

-   At :samp:`<table-row` start: increment row number, reset buffer

-   At :samp:`<table-row` end: yield the row

-   At :samp:`<table-cell` start: check for empty, date, float, boolean types,
    which are available as an attribute at start.
    For strings, start accumulating string values.

-   At :samp:`<table-cell` end: finalize the accumulated value.
