#####################
Schema Module Tests
#####################

A Schema is little more than a list of Attributes,
and an Attribute is little more than a named tuple.

At Attribute, however, is used to build :py:class:`cell.Cell` instances.

See :ref:`schema`.

Overheads
=============

::

  """stingray.schema Unit Tests."""
  import unittest
  import datetime
  import stingray.schema

Mocks
===========

A :py:class:`MockDefaultCell` allows us to make sure that :py:class:`schema.Attribute` builds only this mock object.

::

  class MockDefaultCell:
      def __init__( self, value, workbook ):
          self.value= value
          self.workbook= workbook
      def __eq__( self, other ):
          return all( (
              self.__class__ == other.__class__,
              self.value == other.value,
              ) )
      def __ne__( self, other ):
          return not self.__eq__( other )
        
  class MockNonDefaultCell( MockDefaultCell ):
      pass

  class SchemaMockWorkbook:
      def __init__( self ):
          self.datemode= 0
        
    
Attribute Features
===================

Attributes, generally, have names.  They can have other parameters, but the
attributes  are generally ignored except for COBOL and Fixed format files.

We subclass :py:class:`schema.Attribute` to make it depend on  :py:class:`MockDefaultCell` instead of some working class in :py:mod:`cell`.

::
        
  class AttributeFixture( stingray.schema.Attribute ):
      default_cell= MockDefaultCell

::

  class TestAttribute( unittest.TestCase ):
      def setUp( self ):
          self.simple= AttributeFixture( 
              name="some column" )
          self.complex= AttributeFixture( 
              name="column", offset=0, size=5, create=MockNonDefaultCell )
      def test_should_have_names( self ):
          self.assertEqual( "some column", self.simple.name )
          self.assertEqual( "column", self.complex.name )

Baseline Attribute Conversion 
===============================

Attributes can also have a "create" option
to create a given type of :py:class:`cell.Cell` from input that's not easily decoded.
based on the physical format.

Fixed and COBOL format files are the prime example of cells that require conversions.  A CSV or TAB file could fall into this category, also.

This is an essential feature of the :py:class:`workbook.Fixed_Workbook`.

::

  class TestAttributeConversion( unittest.TestCase ):
      def setUp( self ):
          self.simple= AttributeFixture( 
              name="some column", offset= 0, size= 5 )
          self.complex= AttributeFixture( 
              name="column", offset=5, size=5, create=MockNonDefaultCell )
          self.data = "12345abcde"
      def test_should_extract( self ):
          col= self.simple
          data= self.data
          s= data[col.offset:col.offset+col.size]
          self.assertEqual( "12345", s )
          col= self.complex
          data= self.data
          c= data[col.offset:col.offset+col.size]
          self.assertEqual( "abcde", c )
      def test_should_convert( self ):
          wb= SchemaMockWorkbook()
          col= self.simple
          data= self.data
          s= col.create( data[col.offset:col.offset+col.size], wb )
          self.assertEqual( MockDefaultCell("12345",wb), s )
          col= self.complex
          data= self.data
          c= col.create( data[col.offset:col.offset+col.size], wb )
          self.assertEqual( MockNonDefaultCell("abcde",wb), c )
          self.assertNotEqual( MockDefaultCell("abcde",wb), c )

Schema
===========

A :py:class:`schema.Schema` is essentially a list of attributes.  Order matters, since
names can (and frequently are) duplicates.

::

  class TestSchema( unittest.TestCase ):
      def setUp( self ):
          self.schema = stingray.schema.Schema( 
              AttributeFixture( 
                  name="some column" ),
              AttributeFixture( 
                  name="column", offset=0, size=5, create=MockNonDefaultCell ),
          )
      def test_should_have_names( self ):
          names = [ "some column", "column" ]
          self.assertEqual( names, [ a.name for a in self.schema ] )


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
