##################
Workbook Testing
##################

There are many types of workbooks.  See :ref:`workbook` for details.

These tests use actual files in the :file:`sample` directory.

- :file:`csv_workbook.csv` used by `CSV_Workbook Tests`_.
- :file:`excel97_workbook.xls` used by `XLS_Workbook Tests`_.
- :file:`excel_workbook.xlsx` used by `XLSX_Workbook Tests`_.
- :file:`ooo_workbook.ods` used by `ODS_Workbook Tests`_.
- :file:`tab_workbook.tab` used by `CSV_Workbook with Tab delimiters`_.
- :file:`workbook.simple` used by `Fixed_Workbook Tests`_.  :file:`simple.csv` is the schema.
- :file:`numbers_workbook_09.numbers` used by `Numbers09_Workbook Tests`_.
- :file:`numbers_workbook_13.numbers` used by `Numbers13_Workbook Tests`_.

Overheads
=============

::

  """stingray.workbook Unit Tests."""
  import unittest
  import os
  import decimal
  import datetime
  import stingray.sheet
  import stingray.schema
  import stingray.schema.loader
  import stingray.workbook


CSV_Workbook Tests
=====================

A :py:class:`workbook.base.Workbook`  produces a list of sheet names, is a factory
for individual ``Sheet`` instances and produces the rows of
a named sheet.  Also, a :py:class:`workbook.base.Workbook` is a context.

All subclasses of :py:class:`workbook.base.Workbook` must be polymorphic with this
degenerate special cases.
This means that the tests are nearly identical.  

The most important distinction among the various physical formats
is that other formats (XLS, XLSX, ODS, etc.) support multiple
sheets in a single workbook, where CSV (and Fixed) have only a 
single sheet in the workbook.

::

  class TestCSVWorkbook( unittest.TestCase ):
      theClass = stingray.workbook.CSV_Workbook
      extraKW= {}
      theFile = os.path.join('sample', 'csv_workbook.csv')
      theName = "csv_workbook"
      theSheets = ["csv_workbook"]
      def setUp( self ):
          self.wb = self.theClass( self.theFile, **self.extraKW )
      def test_should_have_sheets( self ):
          self.assertEqual( set(self.theSheets), set(self.wb.sheets()) )
      def test_should_open_sheet_1( self ):
          s = self.wb.sheet( 'Sheet1' )
          row_list= list( s.rows() )
          self.assertEqual( 3, len(row_list) )
          row = row_list[0]
          self.assertEqual( 7, len(row) )
          self.assertEqual( "Col 1 - int", row[0].to_str() )
          self.assertEqual( "Col 2.0 - float", row[1].to_str() )
          self.assertEqual( 'Column "3" - string', row[2].to_str() )
          self.assertEqual( "Column '4' - date", row[3].to_str() )
          self.assertEqual( "Column 5 - boolean", row[4].to_str() )
          self.assertEqual( "Column 6 - empty", row[5].to_str() )
          self.assertEqual( "Column 7 - Error", row[6].to_str() )
      def test_should_have_data_sheet_1( self ):
          s = self.wb.sheet( 'Sheet1' )
          row_list= list( s.rows() )
          self.assertEqual( 3, len(row_list) )
          row = row_list[1]
          self.assertEqual( 42, row[0].to_int() )
          self.assertAlmostEqual( 3.1415926, row[1].to_float() )
          self.assertEqual( 'string', row[2].to_str() )
      def test_should_create_context( self ):
          with self.theClass( self.theFile ) as ctx:
              self.assertEqual( set(self.theSheets), set(ctx.sheets()) )
      def test_should_use_context( self ):
          with self.wb:
              self.assertEqual( set(self.theSheets), set(self.wb.sheets()) )

CSV_Workbook with Tab delimiters
=================================

The idea is that we have a very small change to handle different
delimiters in the ``csv`` module.

::

  class TestTabWorkbook( unittest.TestCase ):
      theClass = stingray.workbook.CSV_Workbook
      extraKW = { 'delimiter':'\t' }
      theFile = os.path.join('sample', 'tab_workbook.tab')
      theName = "tab_workbook"
      theSheets = ["tab_workbook"]


Fixed_Workbook Tests
========================

A :py:class:`workbook.Fixed_Workbook` is similar to a :py:class:`workbook.CSV_Workbook`.
It always has an external schema and the column contents
might be different due to encoding or truncation issues.

::

  class TestFixedWorkbook( TestCSVWorkbook ):
      theClass = stingray.workbook.Fixed_Workbook
      theFile = os.path.join('sample', 'workbook.simple')
      theName = "workbook"
      theSheets = ["workbook"]
      def setUp( self ):
          schema_wb= stingray.workbook.CSV_Workbook( os.path.join( 'sample', 'simple.csv' ) )
          esl = stingray.schema.loader.ExternalSchemaLoader( schema_wb, 
              'sheet1' )
          self.schema= esl.schema()
          self.wb = self.theClass( self.theFile, schema=self.schema )
      def test_should_open_sheet_1( self ):
          s = self.wb.sheet( 'Sheet1' )
          row_list= list( s.rows() )
          self.assertEqual( 3, len(row_list) )
          row = row_list[0]
          self.assertEqual( 7, len(row) )
          self.assertEqual( "Col 1 - int", row[0].to_str() )
          self.assertEqual( "Col 2.0 - f", row[1].to_str() )
          self.assertEqual( 'Column "3"', row[2].to_str() )
          self.assertEqual( "Column '4'", row[3].to_str() )
          self.assertEqual( "Column 5 -", row[4].to_str() )
          self.assertEqual( "Column 6 -", row[5].to_str() )
          self.assertEqual( "Column 7 -", row[6].to_str() )
      def test_should_create_context( self ):
          with self.theClass( self.theFile, schema=self.schema ) as ctx:
              self.assertEqual( set(self.theSheets), set(ctx.sheets()) )

ODS_Workbook Tests
=====================

An ODS workbook is the gold standard.

::

  class TestODSWorkbook( TestCSVWorkbook ):
      theClass = stingray.workbook.ODS_Workbook
      theFile = os.path.join('sample', 'ooo_workbook.ods')
      theName = "ooo_workbook"
      theSheets = ['Sheet1','Sheet2','Sheet3']
      def test_should_open_sheet_2( self ):
          s = self.wb.sheet( 'Sheet2', 
              stingray.sheet.EmbeddedSchemaSheet,
              loader_class= stingray.schema.loader.HeadingRowSchemaLoader )
          row_list= list( s.rows() )
          #print( 'row_list=', row_list )
          self.assertEqual( 3, len(row_list) )
          # First row (after the heading)
          row = row_list[0]
          self.assertEqual( 1, row[0].to_int() )
          self.assertEqual( 'data', row[1].to_str() )
          # Headings
          self.assertEqual( 2, len(s.schema) )
          self.assertEqual( 'Sheet 2 \u2013 int', s.schema[0].name )
          self.assertEqual( 'Sheet 2 \u2013 string', s.schema[1].name )
        
XLS_Workbook Tests
=====================

An XLS workbook can have multiple sheets.  Each sheet can have 
a different schema. This should appear to be like an ODS workbook.

::

  class TestXLSWorkbook( TestODSWorkbook ):
      theClass = stingray.workbook.XLS_Workbook
      theFile = os.path.join('sample', 'excel97_workbook.xls')
      theName = "excel97_workbook"
      theSheets = ['Sheet1','Sheet2','Sheet3']


XLSX_Workbook Tests
=====================

An XLSX workbook should be functionally identical to an ODS workbook.
The physical format is remarkably different, but the content must appear identical.

::

  class TestXLSXWorkbook( TestODSWorkbook ):
      theClass = stingray.workbook.XLSX_Workbook
      theFile = os.path.join('sample', 'excel_workbook.xlsx')
      theName = "excel_workbook"
      theSheets = ['Sheet1','Sheet2','Sheet3']


Numbers09_Workbook Tests
=========================

A Numbers '09 workbook should be functionally similar to an ODS workbook.
Numbers has Workspaces (a/k/a "Pages") with Tables. The list of "sheets" is (workspace,table) pairs.

::

  class TestNumbers09Workbook( TestCSVWorkbook ):
      theClass = stingray.workbook.Numbers09_Workbook
      theFile = os.path.join('sample', 'numbers_workbook_09.numbers')
      theName = "numbers_workbook_09"
      theSheets = [('Sheet 1','Table 1'), ('Sheet 2','Table 1'), ('Sheet 3', 'Table 1')]
      def test_should_open_sheet_1( self ):
          s = self.wb.sheet( ('Sheet 1','Table 1') )
          row_list= list( s.rows() )
          self.assertEqual( 3, len(row_list) )
          row = row_list[0]
          self.assertEqual( 7, len(row) )
          self.assertEqual( "Col 1 - int", row[0].to_str() )
          self.assertEqual( "Col 2.0 - float", row[1].to_str() )
          self.assertEqual( 'Column "3" - string', row[2].to_str() )
          self.assertEqual( "Column '4' - date", row[3].to_str() )
          self.assertEqual( "Column 5 - boolean", row[4].to_str() )
          self.assertEqual( "Column 6 - empty", row[5].to_str() )
          self.assertEqual( "Column 7 - Error", row[6].to_str() )
      def test_should_read_workbook( self ):
          """Read from workbook instead of sheet."""
          s = self.wb.sheet( ('Sheet 1','Table 1') )
          row_list_wb = list( self.wb.rows_of( s ) )
          row_list_s= list( s.rows() )
          self.assertEqual( row_list_wb, row_list_s )
      def test_should_have_data_sheet_1( self ):
          s = self.wb.sheet( ('Sheet 1','Table 1') )
          row_list= list( s.rows() )
          self.assertEqual( 3, len(row_list) )
          row = row_list[1]
          self.assertEqual( 42, row[0].to_int() )
          #print( repr(row[1]) )
          self.assertAlmostEqual( decimal.Decimal('3.1415926'), row[1].to_decimal(7) )
          self.assertAlmostEqual( 3.1415926, row[1].to_float() ) # ?
          self.assertEqual( 'string', row[2].to_str() )
          self.assertEqual( datetime.datetime(1956, 9, 10, 0, 0), row[3].to_datetime() )

Numbers13_Workbook Tests
=========================

A Numbers '13 workbook should be functionally similar to an Numbers '09 workbook.
The format, however, is not XML-based. The content must appear as 
close to identical as practical. 
Numbers has Workspaces (a/k/a "Pages") with Tables. The list of "sheets" is (workspace,table) pairs.

::

  class TestNumbers13Workbook( TestNumbers09Workbook ):
      theClass = stingray.workbook.Numbers13_Workbook
      theFile = os.path.join('sample', 'numbers_workbook_13.numbers')
      theName = "numbers_workbook_13"
      theSheets = [('Sheet 1','Table 1'), ('Sheet 2','Table 1'), ('Sheet 3', 'Table 1')]


Workbook Factory Function
==========================

The :py:func:`workbook.open_workbook()` function should provide 
simple, uniform access to a workbook file.

::

  class TestWBFactory( unittest.TestCase ):
      def test_should_open_csv( self ):
          w = stingray.workbook.open_workbook( os.path.join('sample','csv_workbook.csv') )
          self.assertEqual( set(['csv_workbook']), set(w.sheets()) )
      def test_should_open_fixed( self ):
          w = stingray.workbook.open_workbook(
              os.path.join('sample','workbook.simple'),
              stingray.sheet.ExternalSchemaSheet,
              schema_path='sample',
              schema_sheet='simple', )
          self.assertEqual( set(['workbook']), set(w.sheets()) )
      def test_should_open_xls( self ):
          w = stingray.workbook.open_workbook( os.path.join('sample','excel97_workbook.xls') )
          self.assertEqual( set(['Sheet1','Sheet2','Sheet3']), set(w.sheets()) )
      def test_should_open_xlsx( self ):
          w = stingray.workbook.open_workbook( os.path.join('sample','excel_workbook.xlsx') )
          self.assertEqual( set(['Sheet1','Sheet2','Sheet3']), set(w.sheets()) )
      def test_should_open_ods( self ):
          w = stingray.workbook.open_workbook( os.path.join('sample','ooo_workbook.ods') )
          self.assertEqual( set(['Sheet1','Sheet2','Sheet3']), set(w.sheets()) )
        
        
Test Suite and Runner
=====================

In case we want to build up a larger test suite, we avoid doing
any real work unless this is the main module being executed.

::

  import test
  suite= test.suite_maker( globals() )

  if __name__ == "__main__":
      print( __file__ )
      unittest.TextTestRunner().run(suite())
