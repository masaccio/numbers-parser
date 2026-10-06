..    #!/usr/bin/env python3

..  _`demo_profile`:

##########################################################
Data Profiling Demonstration
##########################################################

This is a simple data profiling application that can be applied
to workbooks to examine the data prior to creating builders.

This produces simple RST-format output on stdout.

A common use case is the following:

..  code-block:: bash

    python3 demo/profile.py sample/\*.csv >profile_csv.rst
    rst2html.py profile_csv.rst profile_csv.html
    
This gives us an HTML-formatted report. 

Overheads
-------------

We depend on a number of Python libraries.  Plus, of course, we're
creating workbooks, working with sheets and schema.

::

    import logging
    import sys
    import argparse
    import pprint
    from collections import defaultdict
    
    import stingray.workbook
    import stingray.sheet
    import stingray.schema
    
    logger= logging.getLogger( __name__ )

Processing Context
---------------------

This class handles the accumulation of global statistics on the source data.
This a context manager which assures that we'll successfully process the 
statistics in spite of any exception which might occur.

::

    class Gather_Statistics:
        stop_on_exception= False
        def __init__( self ):
            self.stats= defaultdict( lambda: defaultdict(int) )
        def count( self, column, value ):
            self.stats[column][value] += 1
        def __enter__( self ):
            return self
        def __exit__( self, exc_type, exc_val, exc_tb ):
            if exc_type is not None: return False
            
See :ref:`demo_validate` for a more complete example of this kind of use
of a context manager.

Processing
------------

See :ref:`developer` for background. We're going to need a "sheet process function."
This transforms the source sheet into the target collection, usually an output file.

The :py:func:`process_sheet` function the heart of the application.
This handles all the rows present in a given sheet.

We use ``source_row[attr]`` to accumulate the :py:class:`cell.Cell` instance
information.  This can help identify the source data format as well as the value.

We can use ``source_row[attr].value`` to accumulate the "raw" Python values
present in the spreadsheet.  

An alternative is to use a :py:class:`cell.Cell` conversion to a desired
type.  We might, for example, use :py:meth:`cell.Cell.to_str` to convert a raw value
to a string.  This would better parallel the way that an application will
use the data.

::

    def process_sheet( sheet, persistence ):
        counts= defaultdict( int )
        for source_row in sheet.schema.rows_as_dict_iter(sheet):
            try:
                counts['read'] += 1
                for attr in source_row:
                    persistence.count( attr, source_row[attr] )
            except Exception as e:
                counts['invalid'] += 1
                if persistence.stop_on_exception: raise
                summary= "{0} '{1}'".format( e.__class__.__name__, e.message )
                logger.error( summary )
                counts['error '+summary] += 1
                
        title= "{0} :: {1}".format( sheet.workbook.name, sheet.name )
        print( title )
        print( "="*len(title) )
        print()
        for attr in sheet.schema:
            name= attr.name
            print( name )
            print( "-"*len(name) )
            print()
            print( "..  csv-table::" )
            print()
            for k in persistence.stats[name]:
                print( '    "{0}","{1}"'.format( k, persistence.stats[name][k] ) )
            print()
        return counts
        
Some applications will have variant processing for workbooks that
contain different types of sheets.  
This leads to different ``process_this_sheet``  and ``process_that_sheet`` functions.  
Each  will follow the above template to process all rows of the sheet.

High-Level Interfaces
------------------------

This version of :py:func:`validate` doesn't really do very much.  We don't have
any persistence, so there's no sensible alternative do this for simple
data gathering. In more complex applications, we might have a :py:func:`process`
function which does some more complex processing.

However, it's often helpful to follow the template design for other,
more sophisticated, applications.  For that reason, we provide the 
processing context as a kind of  **Strategy** object to the :py:func:`process_sheet` function.

::

    def validate( sheet ):        
        with Gather_Statistics() as mode:
            counts= process_sheet( sheet, mode )
        return counts

Note that the following :py:func:`process_workbook` function makes some specific claims about the given
file.  In particular:

-   :py:class:`sheet.EmbeddedSchemaSheet`.  The schema is embedded within each sheet.

-   :py:class:`schema.loader.HeadingRowSchemaLoader`.  The schema is the heading row.

If these assumptions are not universally true, then different application
programs or different :py:func:`process_workbook` functions must be written to handle other kinds of workbooks.

::

    def process_workbook( input ):
        for name in source.sheets():
            logger.info( "{0} :: {1}".format( input, name ) )
            sheet= source.sheet( name, 
                stingray.sheet.EmbeddedSchemaSheet,
                loader_class=stingray.schema.loader.HeadingRowSchemaLoader )
            counts= validate( sheet )
            logger.info( pprint.pformat(dict(counts)) )

Command-Line Interface
----------------------

We have an optional argument for verbosity and a positional argument that
provides all the files to profile.

::

    def parse_args():
        parser= argparse.ArgumentParser()
        parser.add_argument( 'file', nargs='+' )
        parser.add_argument( '-v', '--verbose', dest='verbosity',
            default=logging.INFO, action='store_const', const=logging.DEBUG )
        return parser.parse_args()

The main-import switch allows us to import this module and reuse the components
or run it as a command-line application.  To run from the command line, we have several issues to address.

1.  Logging.
#.  Parameter Parsing.  This includes interpreting options.
#.  Argument Processing.  This means looping over the positional arguments.
#.  Opening Workbooks.  Some applications can't use the default
    :py:class:`workbook.Opener`.  A subclass of Opener, or more complex logic,
    may be required.
#.  Gracefully catching and logging exceptions.
#.  Exit Status to the OS.

::

    if __name__ == "__main__":
        logging.basicConfig( stream=sys.stderr )
        args= parse_args()
        logging.getLogger().setLevel( args.verbosity )
        try:
            for input in args.file:
                with stingray.workbook.open_workbook( input ) as source:
                    process_workbook( source )
            status= 0
        except Exception as e:
            logging.exception( e )
            status= 3 
        logging.shutdown()
        sys.exit( status )
        
Running the Demo
================

We can run this program like this:

..  code-block:: bash

    python3 demo/profile.py sample/\*.csv >profile_csv.rst
    rst2html.py profile_csv.rst profile_csv.html

The RST output file looks like this:

..	parsed-literal::

	sample/csv_workbook.csv :: csv_workbook
	=======================================

	Col 1 - int
	-----------

	..  csv-table::

		"TextCell('9973')","1"
		"TextCell('42')","1"

	Col 2.0 - float
	---------------

	..  csv-table::

		"TextCell('3.1415926')","1"
		"TextCell('2.7182818')","1"

	Column "3" - string
	-------------------

	..  csv-table::

		"TextCell('string')","1"
		"TextCell('data')","1"

	Column '4' - date
	-----------------

	..  csv-table::

		"TextCell('09/10/56')","1"
		"TextCell('01/18/59')","1"

	Column 5 - boolean
	------------------

	..  csv-table::

		"TextCell('TRUE')","1"
		"TextCell('FALSE')","1"

	Column 6 - empty
	----------------

	..  csv-table::

		"TextCell('')","2"

	Column 7 - Error
	----------------

	..  csv-table::

		"TextCell('#DIV/0!')","1"
		"TextCell('#NAME?')","1"

	sample/simple.csv :: simple
	===========================

	name
	----

	..  csv-table::

		"TextCell('Column 6 – empty')","1"
		"TextCell('Column “3” - string')","1"
		"TextCell('Col 2.0 – float')","1"
		"TextCell("Column '4' – date")","1"
		"TextCell('Column 7 – Error')","1"
		"TextCell('Column 5 – boolean')","1"
		"TextCell('Col 1 - int')","1"

	offset
	------

	..  csv-table::

		"TextCell('45')","1"
		"TextCell('56')","1"
		"TextCell('34')","1"
		"TextCell('67')","1"
		"TextCell('1')","1"
		"TextCell('23')","1"
		"TextCell('12')","1"

	size
	----

	..  csv-table::

		"TextCell('11')","7"

	type
	----

	..  csv-table::

		"TextCell('float')","1"
		"TextCell('bool')","1"
		"TextCell('datetime')","1"
		"TextCell('str')","3"
		"TextCell('int')","1"


