..  #!/usr/bin/env python3

.. _`workbook_base`:

Workbook Base Definition
==========================

These are the definitions shared by all Workbook subclass definitions.

One of the common features is date conversions. Only XLS files have 
peculiar date conversion issues. But we're forced to make all 
spreadsheets polymorphic with respect to this anomalous behavior.

..  py:module:: workbook.base

::

  import logging
  import os

  import stingray.sheet

..  py:class:: Workbook

    All physical workbook formats all encode a single, common data structure.
    Here are some abstract definitions.
   
    ..  py:attribute:: name
       
        Filename for this workbook.
   
    ..  py:attribute:: the_file
   
        The actual open file object.
   
    ..  py:attribute:: datemode
   
        For XLS spreadsheets only, a datemode is required. This may not 
        be appropriate for all spreadsheets. It's in this superclass 
        for now, but may be refactored out.
   
    ..  py:attribute:: log
   
        Logger for Workbooks.

::

  class Workbook:
      """A workbook file; a collection of Sheets."""
      def __init__( self, name, file_object=None ):
          """Prepare the workbook for reading.

          :param name: File name
          :param file_object: Optional file-like object.  If omitted, the named file is opened.
          """
          self.name, self.file_obj= name, file_object
          self.the_file = None # Any internal files
          self.datemode= 0 # For xlrd
          self.log= logging.getLogger( self.__class__.__qualname__ )
      def __repr__( self ):
          return "{0}({1!r})".format( self.__class__.__qualname__, self.name )
        
..  py:method:: Workbook.sheet( sheet_name, sheet_type, *args, **kw )

    There are two varieties of sheets, depending on the presence or
    absence of a schema.

::

      def sheet( self, sheet_name, sheet_type=None, *args, **kw ):
          """Returns a :py:class:`sheet.Sheet`, ready for processing."""
          if sheet_type is None: sheet_type= stingray.sheet.Sheet
          sheet = sheet_type( self, sheet_name, *args, **kw )
          return sheet

..  py:method:: Workbook.sheets( )

    The list of sheet names.

::

      def sheets( self ):
          """List of sheet names.
          The filename is a handy default sheet name for CSV and Fixed files.
          """
          nm, _ = os.path.splitext( os.path.basename(self.name) )
          return [ nm ]

The Context Manager interface.

::

      def __enter__( self ):
          return self
      def __exit__( self, exc_type, exc_val, exc_tb ):
          if self.the_file:
              self.the_file.close()
          if exc_type is not None: return False

..  py:method:: Workbook.rows_of( sheet )

    An iterator over rows of a given sheet.

::

      def rows_of( self, sheet ):
          """An iterator over all rows of the given sheet."""
          raise NotImplementedError

..  py:method:: Workbook.row_get( row, attribute )

    Get a value from the current row using the current attribute.
    This is the essential, underlying implementation to fetch data
    from a row of a workbook. 

::

      def row_get( self, row, attribute ):
          """Create a Cell from the row's data."""
          raise NotImplementedError

And, for proper date conversions in XLS spreadsheets only, we have 
two methods. In other spreadsheets, proper system dates and times
are used in somewhat more conventional ways.

..  py:method:: Workbook.float_to_date( value ):

::

      def float_to_date( self, value ):
          raise NotImplementedError

..  py:method:: Workbook.date_to_float( value ):

::

      def date_to_float( value ):
          raise NotImplementedError

There are many distinct subclasses of :py:class:`workbook.base.Workbook`, based on the
physical file format.

Many of our physical formats don't require any physical schema information.
A **Fixed** file, however, requires
a physical format schema definition in order to decompose
each line into cells.
