

.. _`test_cells`:

###################
Cell Module Tests
###################

A Cell is the unit of data.  These unit tests exercise the classes in the
:py:mod:`cell` module.  For more information, see :ref:`cells`.

Overheads
=============

::

  """stingray.cell Unit Tests."""
  import unittest
  import decimal
  import datetime
  import stingray.cell
  import locale

  # The locale varies a bit from version to version
  # and OS to OS.
  SAMPLE_DATE = datetime.datetime(2056, 9, 10).strftime(locale.nl_langinfo(locale.D_FMT))

Cell Tests
===============

Mocks for sheet and workbook to test :py:class:`cell.Cell` features.
A :py:class:`cell.Cell` belongs to a :py:class:`sheet.Sheet` and a :py:class:`sheet.Sheet` belongs
to a :py:class:`workbook.base.Workbook`.  We need a mock for :py:class:`sheet.Sheet` and
:py:class:`workbook.base.Workbook`.

::

  class CellMockWorkbook:
      def __init__( self, mock_date=None, mock_date_float=None ):
          self.datemode= 0
          self.mock_date= mock_date
          self.mock_date_float= mock_date_float
      def date_to_float(self, value):
          return self.mock_date_float
      def float_to_date(self, value):
          if value <= 1: raise ValueError()
          return self.mock_date
        
:py:class:`cell.EmptyCell` is always ``None``.

::

  class TestEmptyCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook()
          self.cell= stingray.cell.EmptyCell('', self.wb)
      def test_should_be_empty( self ):
          self.assertTrue( self.cell.is_empty() )
      def test_should_refuse( self ):
          self.assertIsNone( self.cell.to_str() )
          self.assertIsNone( self.cell.to_int() )
          self.assertIsNone( self.cell.to_float() )
          self.assertIsNone( self.cell.to_decimal() )
          self.assertIsNone( self.cell.to_datetime() )
          self.assertIsNone( self.cell.to_digit_str() )

:py:class:`cell.TextCell` does conversions from text to other forms.  Text
can represent a number of things: numbers, dates or proper
string values.  

In the case of a CSV file, all cells are text cells, and these
conversions are very important.

In the case of XLS files or XLSX files, these conversions are less
important because most cells have sensible type information.

::

  class TestTextCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook()
          self.cell_numb= stingray.cell.TextCell( '123', self.wb )
          self.cell_word= stingray.cell.TextCell( 'xyz', self.wb )
          self.cell_date= stingray.cell.TextCell( SAMPLE_DATE, self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_numb.is_empty() )
          self.assertFalse( self.cell_word.is_empty() )
          self.assertFalse( self.cell_date.is_empty() )
      def test_should_convert_number( self ):
          self.assertEqual( 123, self.cell_numb.to_int() )
          self.assertEqual( 123.0, self.cell_numb.to_float() )
          self.assertEqual( decimal.Decimal('123'), 
              self.cell_numb.to_decimal() )
          self.assertEqual( '00123', self.cell_numb.to_digit_str(5) )
      def test_should_exception_nonnumber( self ):
          with self.assertRaises(ValueError):  
              self.cell_word.to_int()
          with self.assertRaises(ValueError):  
              self.cell_word.to_float()
          with self.assertRaises(decimal.InvalidOperation):  
              self.cell_word.to_decimal()
          with self.assertRaises(ValueError):  
              self.cell_word.to_digit_str()
      def test_should_convert_string( self ):
          self.assertEqual( '123', self.cell_numb.to_str() )
          self.assertEqual( 'xyz', self.cell_word.to_str() )
          self.assertEqual( SAMPLE_DATE, self.cell_date.to_str() )
      def test_should_convert_date( self ):
          self.assertEqual( datetime.datetime(2056,9,10),
              self.cell_date.to_datetime() )
      def test_should_exception_nondate( self ):
          with self.assertRaises(ValueError):  
              self.cell_numb.to_datetime()
          with self.assertRaises(ValueError):  
              self.cell_word.to_datetime()

:py:class:`cell.NumberCell` does some conversions from float to other forms.

::

  class TestNumberCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook(datetime.datetime(1956, 9, 10), 20708)
          self.cell_numb= stingray.cell.NumberCell( 123.4, self.wb )
          self.cell_date= stingray.cell.NumberCell( 20708.0, self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_numb.is_empty() )
          self.assertFalse( self.cell_date.is_empty() )
      def test_should_convert_number( self ):
          self.assertEqual( 123, self.cell_numb.to_int() )
          self.assertEqual( 123.4, self.cell_numb.to_float() )
          self.assertEqual( decimal.Decimal('123.4'), 
              self.cell_numb.to_decimal(1) )
          self.assertEqual( '00123', self.cell_numb.to_digit_str(5) )
          self.assertEqual( '123.4', self.cell_numb.to_str() )
      def test_should_convert_date( self ):
          self.assertEqual( datetime.datetime(1956,9,10), 
              self.cell_date.to_datetime() )

:py:class:`cell.FloatDateCell` does some conversions from a "float that's really a date" to other forms.

::

  class TestFloatDateCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook(datetime.datetime(1956, 9, 10), 20708)
          self.cell_date= stingray.cell.FloatDateCell( 20708.0, self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_date.is_empty() )
      def test_should_convert_number( self ):
          self.assertEqual( 20708, self.cell_date.to_int() )
          self.assertEqual( 20708.0, self.cell_date.to_float() )
          self.assertEqual( decimal.Decimal('20708.0'), 
              self.cell_date.to_decimal(1) )
          self.assertEqual( '20708', self.cell_date.to_digit_str(5) )
          self.assertEqual( '20708.0', self.cell_date.to_str() )
      def test_should_convert_date( self ):
          self.assertEqual( datetime.datetime(1956,9,10), 
              self.cell_date.to_datetime() )

:py:class:`cell.BooleanCell` does some conversions from an "int that's really a true/false"
to other forms.

::

  class TestBooleanCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook(datetime.datetime(1956, 9, 10), 20708)
          self.cell_true= stingray.cell.BooleanCell( 1, self.wb )
          self.cell_false= stingray.cell.BooleanCell( 0, self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_true.is_empty() )
          self.assertFalse( self.cell_false.is_empty() )
      def test_should_convert_number( self ):
          self.assertEqual( 1, self.cell_true.to_int() )
          self.assertEqual( 1.0, self.cell_true.to_float() )
          self.assertEqual( decimal.Decimal('1.0'), 
              self.cell_true.to_decimal(1) )
          self.assertEqual( '00001', self.cell_true.to_digit_str(5) )
          self.assertEqual( '1', self.cell_true.to_str() )
          self.assertEqual( 0, self.cell_false.to_int() )
          self.assertEqual( 0.0, self.cell_false.to_float() )
          self.assertEqual( decimal.Decimal('0.0'), 
              self.cell_false.to_decimal(1) )
          self.assertEqual( '00000', self.cell_false.to_digit_str(5) )
          self.assertEqual( '0', self.cell_false.to_str() )
      def test_should_convert_date( self ):
          with self.assertRaises(ValueError):  
              self.cell_true.to_datetime()
          with self.assertRaises(ValueError):  
              self.cell_false.to_datetime()

:py:class:`cell.ErrorCell` doesn't do many conversions.  Mostly, it raises :py:exc:`ValueError`
when converted to anything other than a string.

::

  class TestErrorCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook(datetime.datetime(1956, 9, 10), 20708)
          self.cell_div0= stingray.cell.ErrorCell( '#DIV/0!', self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_div0.is_empty() )
      def test_should_convert_string( self ):
          self.assertEqual( '#DIV/0!', self.cell_div0.to_str() )
      def test_should_convert_number( self ):
          with self.assertRaises(ValueError):  
              self.cell_div0.to_int()
          with self.assertRaises(ValueError):  
              self.cell_div0.to_float()
          with self.assertRaises(ValueError):  
              self.cell_div0.to_decimal(1)
          with self.assertRaises(ValueError):  
              self.cell_div0.to_digit_str(5)
      def test_should_convert_date( self ):
          with self.assertRaises(ValueError):  
              self.cell_div0.to_datetime()

:py:class:`cell.DateCell` does some conversions from a proper date to other forms.
A proper date doesn't arrive from XLS or XLSX, but can arrive
from CSV or fixed-format sources.

::

  class TestDateCell( unittest.TestCase ):
      def setUp( self ):
          self.wb= CellMockWorkbook(datetime.datetime(1956, 9, 10), 20708)
          now= datetime.datetime(1956, 9, 10, 0, 0, 0)
          self.cell_date= stingray.cell.DateCell( now, self.wb )
      def test_should_be_nonempty( self ):
          self.assertFalse( self.cell_date.is_empty() )
      def test_should_convert_number( self ):
          self.assertEqual( 20708, self.cell_date.to_int() )
          self.assertEqual( 20708.0, self.cell_date.to_float() )
          self.assertEqual( decimal.Decimal('20708.0'), 
              self.cell_date.to_decimal(1) )
          self.assertEqual( '20708', self.cell_date.to_digit_str(5) )
          self.assertEqual( '1956-09-10 00:00:00', self.cell_date.to_str() )
      def test_should_convert_date( self ):
          self.assertEqual( datetime.date(1956,9,10), 
              self.cell_date.to_datetime().date() )

Other Conversions
===================

::

  class TestDateFromString( unittest.TestCase ):
      def test_date_conversion( self ):
          convert= stingray.cell.date_from_string( "%m/%d/%Y" )
          dt= convert( "9/10/1956" )
          self.assertEqual( datetime.datetime(1956, 9, 10), dt )
      def test_datecell_conversion( self ):
          self.wb= CellMockWorkbook( datetime.datetime(1956, 9, 10), 20708 )
          create= stingray.cell.datecell_from_string( "%Y-%m-%d" )
          dc= create( "1956-09-10", self.wb )
          self.assertEqual( datetime.datetime(1956, 9, 10), dc.to_datetime() )
          self.assertEqual( 20708, dc.to_int() )

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
