..  #!/usr/bin/env python3

.. _`workbook_xls`:


XLS Workbook
---------------

::

  import logging
  import pprint

  from stingray.workbook.base import Workbook
  import stingray.sheet
  import stingray.cell

  import xlrd

..  py:module:: workbook.xls

..  py:class:: XLS_Workbook

    Extract sheets, rows and cells from an XLS format file.

    This definition of a workbook wraps :py:mod:`xlrd` so that it fits the Stingray framework.
    We'll use proper :py:class:`cell.Cell` subclass instances instead of the default ``xlrd.Cell``
    values that :py:mod:`xlrd` normally creates.

    In addition to the superclass attributes, some additional unique
    attributes are introduced here.

    ..  py:attribute:: wb

        A xlrd workbook for this file.

    ..  py:attribute:: datemode

        The XLS date mode for this workbook. This is required for converting
        floating-point values to dates and dates to floating-point values.


::

  class XLS_Workbook( Workbook ):
      """Uses ``xlrd``."""
      def __init__( self, name, file_object=None, **kw ):
          """Prepare the workbook for reading.
          :param name: File name
          :param file_object: Optional file-like object.  If omitted, the named file is opened.

          The kw arguments are passed to :py:func:`xlrd.open_workbook`.
          """
          super().__init__( name, file_object )
          if self.file_obj:
              self.wb= xlrd.open_workbook( self.name, file_contents=self.file_obj.read(), **kw )
          else:
              self.wb= xlrd.open_workbook( self.name, **kw )
          self.datemode= self.wb.datemode

..  py:method:: XLS_Workbook.sheets( )

    Return the list of sheets for this workbook.

::

      def sheets( self ):
          """List of sheet names."""
          return self.wb.sheet_names()

We can build an eager :py:class:`sheet.Row` or a :py:class:`sheet.LazyRow` from the available data.
The eager Row includes the conversions.  The LazyRow defers the conversions
until the callback to :py:meth:`XLS_Workbook.row_get`.

..  py:method:: XLS_Workbook.rows_of( sheet )

    Iterate through rows of the given sheet.


::

      def rows_of( self, sheet ):
          """An iterator over all rows of the given sheet object.
          This uses the XLRD translation to locate the sheet.
          """
          current_sheet= self.wb.sheet_by_name(sheet.name)
          for n in range(current_sheet.nrows):
              data = current_sheet.row(n)
              row = stingray.sheet.Row( sheet, *(self.cell(col) for col in data) )
              yield row

..  py:method:: XLS_Workbook.row_get( row, attribute )

    Low-level get of a particular attribute from the given row.

::

      def row_get( self, row, attribute ):
          """Create a Cell from the row's data."""
          return row[attribute.position]

..  py:method:: XLS_Workbook.cell( row, xlrd_cell )

    In :py:meth:`XLS_Workbook.rows_of` we built a row eagerly.
    That way, returning an individual Cell is easy.

    Convert a single ``xlrd.Cell`` to a proper subclass of :py:class:`cell.Cell`

::

      def cell( self, xlrd_cell ):
          if xlrd_cell.ctype == xlrd.XL_CELL_EMPTY:
              return stingray.cell.EmptyCell('', self)
          elif xlrd_cell.ctype == xlrd.XL_CELL_TEXT:
              return stingray.cell.TextCell( xlrd_cell.value, self )
          elif xlrd_cell.ctype == xlrd.XL_CELL_NUMBER:
              return stingray.cell.NumberCell( xlrd_cell.value, self )
          elif xlrd_cell.ctype == xlrd.XL_CELL_DATE:
              return stingray.cell.FloatDateCell( xlrd_cell.value, self )
          elif xlrd_cell.ctype == xlrd.XL_CELL_BOOLEAN:
              return stingray.cell.BooleanCell( xlrd_cell.value, self )
          elif xlrd_cell.ctype == xlrd.XL_CELL_ERROR:
              return stingray.cell.ErrorCell(
                  xlrd.error_text_from_code[xlrd_cell.value], self )
          elif xlrd_cell.ctype == xlrd.XL_CELL_BLANK:
              return stingray.cell.EmptyCell('', self)
          else:
              raise ValueError( "Damaged Workbook" )

For proper date conversions, we have
two methods that leverage the datemode to properly convert dates
and times in :file:`.XLS` workbooks.

..  py:method:: Workbook.float_to_date( value ):

::

      def float_to_date( self, value ):
          try:
              dt= xlrd.xldate_as_tuple(value, self.datemode)
          except xlrd.xldate.XLDateAmbiguous as e:
              ex= ValueError( "Ambiguous Date: {0!r}".format(value) )
              raise ex from e
          return datetime.datetime(*dt)

..  py:method:: Workbook.date_to_float( value ):

::

      def date_to_float( value ):
          timetuple= self.value.timetuple()[:6]
          xl= xlrd.xldate.xldate_from_datetime_tuple(
              timetuple,
              self.datemode)
          return xl
