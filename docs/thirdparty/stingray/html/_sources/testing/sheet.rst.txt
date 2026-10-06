
###################
Sheet Module Tests
###################

A Sheet is a collection of Cells, with a structure imposed by a Schema.
These unit tests exercise the classes in the
:py:mod:`sheet` module.  For more information, see :ref:`sheets`.

Overheads
=============

::

  """stingray.sheet Unit Tests."""
  import unittest
  import decimal
  import datetime
  import stingray.cell
  import stingray.sheet

Sheet Tests
==============

The top-level :py:class:`sheet.Sheet` simply produces row-as-list from a workbook.

First, we define a :py:class`MockWorkbook` class to implement a minimal interface
that a :py:class:`sheet.Sheet` can rely on.

::

  class MockWorkbook:
      def rows_of( self, sheet ):
          self.requested= sheet.name
          for r in self.rows:
              yield stingray.sheet.Row( sheet, *r )
      def sheet( self, name ):
          return self.mock_sheet

Given that, we can define a sensible unit test.

::

  class TestSheet( unittest.TestCase ):
      def setUp( self ):
          self.wb= MockWorkbook( )
          self.sheet= stingray.sheet.Sheet( self.wb, 'The_Name' )
          self.wb.rows = [
              [ stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb) ],
              [ stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ],
              ]
      def test_should_iterate( self ):
          row_list= list( self.sheet.rows() )
          self.assertEqual( 'The_Name', self.wb.requested )
          self.assertEqual( 2, len(row_list) )
          self.assertTrue( all( isinstance(r,stingray.sheet.Row) for r in row_list ) )
          row= row_list[0]
          self.assertEqual( 1, row[0].to_int() )
          self.assertEqual( "First", row[1].to_str() )
          row= row_list[1]
          self.assertEqual( 2, row[0].to_int() )
          self.assertEqual( "Second", row[1].to_str() )
      def test_should_have_attributes( self ):
          self.assertEqual( 'The_Name', self.sheet.name )

Eager Row Tests
==================

An eager row is just a tuple, created by a workbook when requested by a sheet.

::

  class MockSheet:
      def __init__( self, workbook, name, schema ):
          self.workbook= workbook
          self.name= name
          self.schema= schema
      def rows( self ):
          return self.workbook.rows_of( self )

An eager Row can be built by many of the worksheets where the physical
format provides guidance on field structure and data type conversions.

::

  class TestEagerRow( unittest.TestCase ):
      def setUp( self ):
          self.wb= MockWorkbook( )
          self.sheet= MockSheet( self.wb, 'The_Name', None )
          self.wb.rows = [
              ( stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb), ),
              ( stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ),
              ]
      def test_should_iterate_eager( self ):
          row_list= list( self.sheet.rows() )
          self.assertEqual( 'The_Name', self.wb.requested )
          self.assertEqual( 2, len(row_list) )
          self.assertTrue( all( isinstance(r,stingray.sheet.Row) for r in row_list ) )
          row= row_list[0]
          self.assertEqual( 1, row[0].to_int() )
          self.assertEqual( "First", row[1].to_str() )
          row= row_list[1]
          self.assertEqual( 2, row[0].to_int() )
          self.assertEqual( "Second", row[1].to_str() )            

Lazy Row Tests
===================

An Lazy Row must be built by Fixed format and COBOL format. The physical
format provides zero guidance on field structure and data type conversions.

::

  class MockLazyWorkbook:
      def rows_of( self, sheet ):
          self.requested= sheet.name
          for r in self.rows:
              yield stingray.sheet.LazyRow( sheet, data=r )
      def sheet( self, name ):
          return self.mock_sheet
      def row_get( self, row, attribute ):
          return row._state['data'][attribute.position]

::

  class MockSchema(list):
      def __init__( self, *args, **kw ):
          super().__init__( args )
          self.info= kw
        
::

  class MockAttribute:
      def __init__( self, **kw ):
          self.__dict__.update( kw )
        
::

  class TestLazyRow( unittest.TestCase ):
      def setUp( self ):
          self.schema= MockSchema(
              MockAttribute( name="col1", position=0, signed=False, ),
              MockAttribute( name="col2", position=1, signed=False, ),
              dde=[]
          )
          self.wb= MockLazyWorkbook( )
          self.sheet= MockSheet( self.wb, 'The_Name', self.schema )
          self.wb.rows = [
              ( stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb), ),
              ( stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ),
              ]
      def test_should_iterate_lazy( self ):
          row_list= list( self.sheet.rows() )
          self.assertEqual( 'The_Name', self.wb.requested )
          self.assertEqual( 2, len(row_list) )
          self.assertTrue( all( isinstance(r,stingray.sheet.LazyRow) for r in row_list ) )
          row= row_list[0]
          self.assertEqual( 1, row[0].to_int() )
          self.assertEqual( "First", row[1].to_str() )
          row= row_list[1]
          self.assertEqual( 2, row[0].to_int() )
          self.assertEqual( "Second", row[1].to_str() )            

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
