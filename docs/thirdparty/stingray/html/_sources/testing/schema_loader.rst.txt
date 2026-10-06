######################
Schema Loader Testing
######################


A Schema Loader will build a schema from a variety of sources.

-   Embedded in the first row of a sheet.  
    The :py:class:`schema.loader.HeadingRowSchemaLoader` handles this.

-   Embedded elsewhere in a sheet.  
    A subclass of :py:class:`schema.loader.SchemaLoader` can handle this.

-   An external schema, read from a separate file.  
    The :py:class:`schema.loader.ExternalSchemaLoader` or a subclass can handle this.

-   A "hard-coded" schema, written as a Python object definition.  
    A static instance of :py:class:`schema.Schema` handles this gracefully.
   
For the essential implementation, see :ref:`schema_loader`.

Also, we will have schemata written in COBOL notation, but that's handled separately.

Overheads
=============

::

  """stingray.schema.loader Unit Tests."""
  import unittest
  import stingray.schema
  import stingray.sheet
  import stingray.schema.loader


First, we define mock classes.  

:py:class:`MockWorkbook` class  implements a minimal interface that a :py:class:`sheet.Sheet` can rely on.

::

  class SheetMockWorkbook:
      def rows_of( self, sheet ):
          self.requested= sheet.name
          return iter( self.rows )
      def sheet( self, name, sheet_type, **kw ):
          return self.mock_sheet

:py:class:`MockSheet` class  implements a minimal interface that a :py:class:`sheet.Row` can rely on.

::

  class MockSheet:
      def __init__( self, workbook, sheet_name ):
          self.workbook, self.name= workbook, sheet_name
      def schema( self ):
          return self.mock_schema
      def rows( self ):
          return self.workbook.rows_of( self )

:py:class:`MockRow` class  implements a minimal interface that collects cell instances.

::

  class MockRow:
      def __init__( self, sheet, *data ):
          self.sheet= sheet
          self.data= data
      def cell( self, attribute ):
          return self.data[attribute.position]
      def __iter__( self ):
          return iter(self.data)
        
HeadingRowSchemaLoader Tests
==============================

A :py:class:`schema.loader.SchemaLoader`` is abstract, and doesn't require much independent testing.

A :py:class:`schema.loader.HeadingRowSchemaLoader` builds a :py:class:`schema.Schema` from the first row from a :py:class:`sheet.Sheet`.

::

  class TestHeadingRowSchemaParser( unittest.TestCase ):
      def setUp( self ):
          self.wb= SheetMockWorkbook( )
          self.sheet= MockSheet( self.wb, 'The_Name' )
          self.wb.rows = [
              MockRow( self.sheet, stingray.cell.TextCell("Col 1", self.wb), 
                stingray.cell.TextCell("Col 2", self.wb) ),
              MockRow( self.sheet, stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb) ),
              MockRow( self.sheet, stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ),
              ]
          self.parser= stingray.schema.loader.HeadingRowSchemaLoader( self.sheet )
      def test_should_parse( self ):
          schema = self.parser.schema()
          self.assertEqual( 2, len(schema) )
          self.assertEqual( "Col 1", schema[0].name )
          self.assertEqual( "Col 2", schema[1].name )
          rows = list( self.parser.rows() )
          self.assertEqual( 2, len(rows) )


ExternalSchemaLoader
=======================

An :py:class:`schema.loader.ExternalSchemaLoader`  builds
a :py:class:`schema.Schema`  from another :py:class:`workbook.base.Workbook`.   

::

  class TestExternalSchemaLoader( unittest.TestCase ):
      def setUp( self ):
          self.wb= SheetMockWorkbook( )
          self.wb.mock_sheet= MockSheet( self.wb, 'The_Name' )
          self.wb.mock_sheet.schema = stingray.schema.Schema( 
              stingray.schema.Attribute("name",create=stingray.cell.TextCell),
              stingray.schema.Attribute("offset",create=stingray.cell.NumberCell),
              stingray.schema.Attribute("size",create=stingray.cell.NumberCell),
              stingray.schema.Attribute("type",create=stingray.cell.TextCell),
          )
          sheet= self.wb.mock_sheet
          self.wb.rows = [
              #MockRow( sheet, stingray.cell.TextCell("name", self.wb), 
              #  stingray.cell.TextCell("offset", self.wb),
              #  stingray.cell.TextCell("size", self.wb),
              #  stingray.cell.TextCell("type", self.wb),
              #  ),
              MockRow( sheet, stingray.cell.TextCell("Col 1", self.wb),
                stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.NumberCell(5, self.wb), 
                stingray.cell.TextCell("str", self.wb ),
                ),
              MockRow( sheet, stingray.cell.TextCell("Col 2", self.wb),
                stingray.cell.NumberCell(6, self.wb), 
                stingray.cell.NumberCell(10, self.wb), 
                stingray.cell.TextCell("str", self.wb),
                ),
              ]
            
          self.loader= stingray.schema.loader.ExternalSchemaLoader( self.wb, 'The_Name',)
      def test_should_parse( self ):
          schema = self.loader.schema()
          self.assertEqual( 2, len(schema) )
          self.assertEqual( "Col 1", schema[0].name )
          self.assertEqual( "Col 2", schema[1].name )


ExternalSchemaSheet Tests
============================

The :py:class:`sheet.ExternalSchemaSheet` creates a schema from an external 
source, usually another Workbook.

An external schema can be loaded using an :py:class:`schema.loader.ExternalSchemaLoader`.
However, we can also hard-code a Schema.  Or build it some other way.

A :py:class:`workbook.base.Workbook` uses this type to build each individual :py:class:`sheet.Sheet`,
attaching the schema (and column titles, etc.) to each :py:class:`sheet.Row`
that gets built.

::

  class TestExternalSchemaSheet( unittest.TestCase ):
      def setUp( self ):
          self.wb= SheetMockWorkbook( )
          schema = stingray.schema.Schema( 
              stingray.schema.Attribute( "Col 1" ),
              stingray.schema.Attribute( "Col 2" ),
          )
          self.sheet= stingray.sheet.ExternalSchemaSheet( self.wb, 'The_Name', schema=schema )
          self.wb.rows = [
              MockRow( self.sheet, stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb) ),
              MockRow( self.sheet, stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ),
              ]
      def test_should_iterate( self ):
          rows_as_dict_iter = ( 
              dict( zip( (a.name for a in self.sheet.schema), r ) ) 
              for r in self.sheet.rows() )
          row_list= list(  rows_as_dict_iter )
          self.assertEqual( 'The_Name', self.wb.requested )
          self.assertEqual( 2, len(row_list) )
          self.assertTrue( all( isinstance(r,dict) for r in row_list ) )
          row= row_list[0]
          self.assertEqual( 1, row['Col 1'].to_int() )
          self.assertEqual( "First", row['Col 2'].to_str() )
          row= row_list[1]
          self.assertEqual( 2, row['Col 1'].to_int() )
          self.assertEqual( "Second", row['Col 2'].to_str() )
      def test_should_have_attributes( self ):
          self.assertEqual( 'The_Name', self.sheet.name )
        
EmbeddedSchemaSheet Tests
============================

An :py:class:`sheet.EmbeddedSchemaSheet` creates a schema from the workbook and
produces rows.  It relies on a :py:class:`schema.loader.HeadingRowSchemaLoader`, which 
we must mock.

Here's a mock schema loader. 

:: 

  class MockLoader:
      def __init__( self, sheet ):
          self.sheet= sheet
      def schema( self ):
          schema= stingray.schema.Schema( 
              stingray.schema.Attribute( "Col 1" ),
              stingray.schema.Attribute( "Col 2" )
              )
          return schema
      def rows( self ):
          row_iter= iter( self.sheet.rows() )
          skip_this= next(row_iter)
          return row_iter

::

  class TestEmbeddedSchemaSheet( unittest.TestCase ):
      def setUp( self ):
          self.wb= SheetMockWorkbook( )
          self.sheet= stingray.sheet.EmbeddedSchemaSheet( self.wb, 'The_Name', loader_class=MockLoader )
          self.wb.rows = [
              MockRow( self.sheet, stingray.cell.TextCell("Col 1", self.wb), 
                stingray.cell.TextCell("Col 2", self.wb) ),
              MockRow( self.sheet, stingray.cell.NumberCell(1, self.wb), 
                stingray.cell.TextCell("First", self.wb) ),
              MockRow( self.sheet, stingray.cell.NumberCell(2, self.wb), 
                stingray.cell.TextCell("Second", self.wb), ),
              ]
      def test_should_iterate( self ):
          rows_as_dict_iter = ( 
              dict( zip( (a.name for a in self.sheet.schema), r ) ) 
              for r in self.sheet.rows() )
          row_list= list(  rows_as_dict_iter )
          self.assertEqual( 'The_Name', self.wb.requested )
          self.assertEqual( 2, len(row_list) )
          self.assertTrue( all( isinstance(r,dict) for r in row_list ) )
          row= row_list[0]
          self.assertEqual( 1, row['Col 1'].to_int() )
          self.assertEqual( "First", row['Col 2'].to_str() )
          row= row_list[1]
          self.assertEqual( 2, row['Col 1'].to_int() )
          self.assertEqual( "Second", row['Col 2'].to_str() )
      def test_should_have_attributes( self ):
          self.assertEqual( 'The_Name', self.sheet.name )


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
