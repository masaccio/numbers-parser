..    #!/usr/bin/env python3

..  _`demo_sqa`:

#######################################################
Unit Level Validation for Application and Data
#######################################################

We validate that a given file actually matches the required schema through a three-part valiation process.  

1.  Validate application's use of a schema via conventional unit testing.
    This is `Unit Test The Builder Function`_, the first part of our testing.

2.  Validate file conformance to a schema via "live-file testing".
    This is `Live Data Test The Builder Function`_, the second part of our testing. 

3.  Validate the three-way application-schema-file binding by including a
    **File Validation** mode in every file processing application.  
    We'll look at the part 3 (the 3-way binding of app-schema-file) in :ref:`demo_validate`.

Of course, we must do good unit testing on the application overall.  
We'll assume that without further discussion.  Failure to test is simply failure.

Live file testing moves beyond simple unit testing into a realm of establishing evidence
that an application will process a given file correctly.  We're gathering auditable
historical information on file formats and contents. 

Our focus here is validation of the application-schema binding as well as the file-schema binding.  

We can use the  :py:mod:`unittest` module to write tests that 
validate the schema shared by an application and file.  
There are three levels to this validation.

-   Unit-level.  This is a test of the builder functions more-or-less in isolation.  
    There are two kinds of builder validation.
    
    -   Subsets of rows.  Developers usually start here.
    
    -   Whole ("live") files.  Production files are famous for having a few "odd" rows.  
        Whole-file tests are essential.  
    
-   Integration-level.  This is a test of a complete application using
    A test database (or collection of files) is prepared and the application 
    actually exercised against actual files.  
        
Not *everything* is mocked here: the "unit" is an integrated collection of 
components. Not an isolated class.

Unit level tests look like this.

Overheads
==========

::

    import unittest
    from collections import defaultdict
    import stingray.cell
    import stingray.sheet
    import stingray.workbook
    import stingray.schema
    import stingray.schema.loader

Builder Functions
==================

See :ref:`developer` for background. We're going to need a "builder function."
This transforms the source row object into the target object or collection.

Normally, we import the unit under test.
It would look like this:
    
..  parsed-literal::

    from demo.some_app import some_builder

For this demo, here's a sample builder function:

..  py:function:: some_builder(aRow)

    Build a Python dict from the row of data.

::

    def some_builder( aRow ):
        return dict( 
            key= aRow['Column "3" - string'].to_str(),
            value= aRow['Col 2.0 - float'].to_float()
        )

This depends on a schema that permits eager row building. That's 
common outside COBOL file processing.

Mock Workbook
===============

We'll use a mock :py:class:`Workbook` to slightly simplify the builder-in-isolation test.

::

    class MockWorkbook:
        def rows_of( self, sheet ):
            self.requested= sheet.name
            for r in self.mock_rows:
                yield stingray.sheet.Row( sheet, *r )
        def sheet( self, name ):
            return mock_sheet
        def row_get( self, row, attr ):
            return row[attr.position]

Unit Test The Builder Function
===============================
            
Step one is to unit test the :py:func:`some_builder` function with selected row-like objects
from a the :py:class:`MockWorkbook`.

There are two parts: a :py:class:`schema.Schema` and 
a :py:class:`sheet.ExternalSchemaSheet` that contains the mock row(s).

::

    class Test_Builder_2( unittest.TestCase ):
        def setUp( self ):
            self.schema= stingray.schema.Schema( 
                stingray.schema.Attribute( name='Col 1 - int' ), 
                stingray.schema.Attribute( name='Col 2.0 - float' ), 
                stingray.schema.Attribute( name='Column "3" - string' ), 
                stingray.schema.Attribute( name="Column '4' - date" ), 
                stingray.schema.Attribute( name='Column 5 - boolean' ), 
                stingray.schema.Attribute( name='Column 6 - empty' ), 
                stingray.schema.Attribute( name='Column 7 - Error' ), )
            self.wb= MockWorkbook()
            self.sheet= stingray.sheet.ExternalSchemaSheet( self.wb, "Test", self.schema )
            self.row= [
                stingray.cell.NumberCell(42.0, self.wb), 
                stingray.cell.NumberCell(3.1415926, self.wb), 
                stingray.cell.TextCell('string', self.wb), 
                stingray.cell.FloatDateCell(20708.0, self.wb), 
                stingray.cell.BooleanCell(1, self.wb),
                stingray.cell.EmptyCell(None, self.wb),
                stingray.cell.ErrorCell('#DIV/0!'), ]
            self.wb.mock_sheet= self.sheet
            self.wb.mock_rows= [ self.row, ]
        def test_should_build_from_row( self ):
            row= next( self.sheet.rows() )
            dict_row= dict( (a.name, row.cell(a)) for a in self.schema )
            result= some_builder( dict_row )
            self.assertEqual( 'string', result['key'] )
            self.assertAlmostEqual( 3.1415926, result['value'] )

Live Data Test The Builder Function
====================================

Step two is to unit test the :py:func:`some_builder` function with all rows in a given workbook.
In this demo, we're using :file:`sample/excel97_workbook.xls`. Generally, we want to compute
some aggregate (like a checksum) of various data items to be sure we've read and 
converted them properly. 

Pragmatically, it's sometimes hard to get a proper checksum, so we have to resort
to sums, counts, and perhaps even frequency distribution tables.

::

    class Test_Builder_2_Live( unittest.TestCase ):
        def setUp( self ):
            self.wb= stingray.workbook.open_workbook( "sample/excel97_workbook.xls" )
            self.sheet = stingray.sheet.EmbeddedSchemaSheet( 
                self.wb, 'Sheet1', 
                stingray.schema.loader.HeadingRowSchemaLoader )
            self.schema= self.sheet.schema
        def test_should_build_all_rows( self ):
            summary= defaultdict( int )
            for row in self.sheet.rows():
                dict_row= dict( (a.name,row.cell(a)) for a in self.schema )
                result= some_builder( dict_row )
                summary[result['key']] += 1
            self.assertEqual( 1, summary['string'] )
            self.assertEqual( 1, summary['data'] )

Sheet-Level Testing
========================
        
See :ref:`developer` for background. We're going to need a "sheet process function."
This transforms the source sheet into the target collection, usually an output file.

We can also use the unit test tools to test the application's
higher-level :py:func:`process_some_sheet` function.

Normally, we import the unit under test.
It looks like this:
    
..  parsed-literal::

    from demo.some_app import process_some_sheet

For this demo, here's a sample sheet process function:

..  py:function:: process_some_sheet(sheet)

    Process all rows of the named sheet.

::

    def process_some_sheet( sheet ):
        counts= defaultdict( int )
        for row in sheet.schema.rows_as_dict_iter(sheet):
            row_dict= some_builder( row )
            counts['key',row_dict['key']] += 1
            counts['read'] += 1
        return counts

The unit test checks the embedded schema and the overall row counts
from processing the live file.

In this demo, we're using :file:`sample/excel97_workbook.xls`. 

The test opens the workbook. It selects a sheet from the workbook using the class
that extracts the schema from the row headers. The test then uses the :py:func:`process_some_sheet`
function on the given sheet to extract data. In this case, the extraction is 
a frequency table.

::

    class Test_Sheet_Builder_2_Live( unittest.TestCase ):
        def setUp( self ):
            self.wb= stingray.workbook.open_workbook( "sample/excel97_workbook.xls" )
            self.sheet = stingray.sheet.EmbeddedSchemaSheet( 
                self.wb, 'Sheet1', 
                loader_class=stingray.schema.loader.HeadingRowSchemaLoader )
        def test_should_load_schema( self ):
            self.assertEqual( 'Col 1 - int', self.sheet.schema[0].name )
            self.assertEqual( 'Col 2.0 - float', self.sheet.schema[1].name )
            self.assertEqual( 'Column "3" - string', self.sheet.schema[2].name )
            self.assertEqual( "Column '4' - date", self.sheet.schema[3].name )
            self.assertEqual( 'Column 5 - boolean', self.sheet.schema[4].name )
            self.assertEqual( 'Column 6 - empty', self.sheet.schema[5].name )
            self.assertEqual( 'Column 7 - Error', self.sheet.schema[6].name )
        def test_should_build_sample_row( self ):
            counts= process_some_sheet( self.sheet )
            self.assertEqual( 2, counts['read'] )
            self.assertEqual( 1, counts['key','string'] )
            self.assertEqual( 1, counts['key','data'] )

Main Program Switch
====================

This is a common unittest main program.  Ideally, we'd create an actual
suite object to allow combining tests.

::

    if __name__ == "__main__":
        unittest.main()
        
Running the Demo
=================

We can run this program like this:

..  code-block:: bash

    python3 demo/test.py

This produces ordinary unitest output.

..  parsed-literal::

    ....
    ----------------------------------------------------------------------
    Ran 4 tests in 0.016s

    OK
