..    #!/usr/bin/env python3

..  _`demo_validate`:

##########################################################
Application Level Data Validation Technique
##########################################################

We validate that that a file actually matches a schema through a three-part valiation process.  
We looked at the first two parts in :ref:`demo_sqa`.

-   Validate an application's use of a schema via conventional unit testing.

-   Validate file conformance to a schema via "live-file testing".

In this section we'll show how to include 
a **File Validation** mode in every file processing application.  

Having a validation mode means that we must disentangle all of the persistent state change 
operations from the input and processing in our application.  
The "normal" mode uses persistent changes based on the output.
The validation mode doesn't make persistent changes;
it can be viewed as a sort of "dry run": all the processing; none of the writing.

We'll do this with a combination of the **Command** and the **Strategy** design patterns.
We'll create applications
which validate their input file and have a simple plug-in strategy for 
doing any final persistent processing on that file.

..  note:: Simple File Structures

    This validation is designed for simple CSV files with embedded schema.
    The assumption is that each sheet within the workbook
    has a consistent structure. There's no filter applied to pass
    or reject sheets.
            
    Some kind of extension to this application is required
    to handle named sheets within a more complex workbook or 
    to handle sheets which have no header.

State Change Commands
=======================

The **Command** design patterns is helpful for isolating state changes in an application.

Each change (create, update, delete) creates a **Command**.  
In validate mode, these are created but not applied. 
In "normal" processing mode, these are created and applied.

For Extract-Transform-Load (ETL) applications, the commands are the loads.

For create-retrieve-update-delete (CRUD) programs, the commands are variations on create, update and delete.

For data warehouse dimensional conformance applications, the 
command may be a slowly-changing dimension (SCD) algorithm that does insert
or update (or both) into a dimension table.

For applications that involve a (potentially) complex multi-step workflow with
(potentially) several state changes along the way, each change is a command.

In some cases, a fairly sophisticated **Command** class hierarchy is
called for.  In other cases, however, the individual commands can be
merged into the validate **Strategy** object as methods.

Persistence Context Manager
=============================

One good way to distinguish between persistent and transient processing 
is to use a **Strategy** class hierarchy.
This will have two variations on the persistent state changes.

-   **Validate**.  This subclass does nothing.

-   **Process**.  This subclass actually makes persistent state changes.

Combining the validate **Strategy** with the state change **Command** 
leads to class similar to the following.

The superclass does the persistent processing. This is the "normal" mode
that makes proper changes to the filesystem or database.

..  parsed-literal::

    class Persistent_Processing:
        stop_on_exception= True
        def __init__( self, context ):
            self.context= context
        def save_this( self, this_instance ):
            this_instance.save()
        def save_that( self, this_instance ):
            that_instance.save()

We'll fold in the Context Manager interface.  This is a polite way to support
any preparation or finalization.  For example, we would use the context manager
to create database connections, or finalize file system operations, etc.

..  parsed-literal::

        def __enter__( self ):
            return self
        def __exit__( self, exc_type, exc_val, exc_tb ):
            if exc_type is not None: return False
     
Here's a subclass which implements a safe, do-nothing strategy.  This 
is used for "validate-mode" processing. It's a subclass that turns off
persistence.
       
..  parsed-literal::

    class Validate_Only_Processing( Persistent_Processing ):
        stop_on_exception= False
        def __init__( self, context ):
            self.context= context
        def save_this( self, this_instance ):
            pass
        def save_that( self, this_instance ):
            pass

..	note:: Alternate Design

	We could revise this design to make the validation mode the superclass.
	The subclass could then add the persistence features.
	
	This doesn't actually work out well in practice.
	
	Why not?
	
	It's too easy to overlook things in the validation mode superclass.
	The normal persistent processing subclass then winds up having a **lot** of extra 
	stuff added to it.
	
	The design winds up somewhat better looking if we remove persistence.
	
Having these two classes allows us to configure our 
application processing as follows. We can define high-level functions
like :py:func:`validate` and :py:func:`process` that are identical
except for the context manager that's used.

..  parsed-literal::

    def validate( sheet, some_context ):
        with Validate_Only_Processing( some_context ) as mode:
            counts= process_sheet( sheet, mode )
        return counts

    def process( sheet, some_context ):        
        with Persistent_Processing( some_context ) as mode:
            counts= process_sheet( sheet, mode )
        return counts

Both of these :py:func:`validate` and :py:func:`process` functions
rely on a common :py:func:`process_sheeet`. This is agnostic of the
processing context; it simply does its work.

..  parsed-literal::

    def process_sheet( sheet, persistence ):
        for row in sheet.schema.rows_as_dict_iter(sheet):
            try:
                this= build_this( row )
                f= ThisForm( this )
                if f.is_valid():
                    persistence.save_this( this )
            except Exception, e:
                if persistence.stop_on_exception: raise
                
This allows us to effectively unit test by creating a 
mock version of ``Persistent_Processing`` and invoking
the ``process_sheet`` function.

Example Application
=======================

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

ORM Layer
---------

We'll often have an Object-Relational Mapping (ORM) layer.
These are components that are widely shared.  They could be SQLAlchemy 
mapped class or a Django ORM class.  

It's appropriate to supplement the ORM with a "Form" that
follows the Django design pattern. This is a class that is used for validating 
and creating model instances.

Here's our fake model object, :py:class:`This`, and it's form, :py:class:`ThisForm`.

::

    class This:
        def __init__( self, key, value ):
            self.key, self.value= key, value
        def save( self ):
            pass # The ORM save operation
            
    class ThisForm:
        def __init__( self, **kw ):
            self.clean= kw
        def is_valid( self ):
            return self.clean['key'] is not None
        def save( self ):
            return This( **self.clean )

We need to be sure that the ORM's save operation is only invoked through our persistence
processing **Strategy** object.  With some libraries the persistence can 
be implicit, making it difficult to assure that persistence is disabled properly.  

There are a number of ways to handle implicit persistence in ORM layers.
It may be necessary to provide a mock database "engine" or interface 
in order to disable persistence.

Persistence Context Manager
----------------------------

These two classes define our two modes: validation and normal operations.
The superclass defines the normal processing mode: we actually save objects.
The subclass defines validation-only mode: we don't save anything.

::

    class Persistent_Processing:
        stop_on_exception= True
        def save_this( self, this_instance ):
            this_instance.save()
        def __enter__( self ):
            return self
        def __exit__( self, exc_type, exc_val, exc_tb ):
            if exc_type is not None: return False
                            
    class Validate_Only_Processing( Persistent_Processing ):
        stop_on_exception= False
        def save_this( self, this_instance ):
            pass

In larger and more sophisticated applications, there may be a much more
complex set of class definitions to enable or disable persistence.

Builder Functions
-----------------

See :ref:`developer` for background. We're going to need a "builder function."
This transforms the source row object into the target object or collection.

To handle variant logical layouts, a number of builder functions are provided
to map the logical schemata to a more standardized conceptual schema.

::
        
    def builder_1( row ):
        return dict( 
            key= row['Column "3" - string'].to_str(),
            value= row['Col 2.0 - float'].to_float()
        )
        
    def builder_2( row ):
        return dict( 
            key= row['Column 3'].to_str(),
            value= row['Column 2'].to_float()
        )
        
Note that these builder functions are frequently added to. It's rare to get these 
down to a single version that always works.

Consequently, it's important to always **add** new builder functions.  Logical layouts are a
moving target.  Old layouts don't go away; making changes to a builder is a bad idea.

It helps to have a function like this to map argument values to a builder function.

::

    def make_builder( args ):
    	return {
    		'1': builder_1,
    		'2': builder_2
    		}[args.layout]

It can help to have a better naming convention that "_1" and "_2".  In practice,
however, it's sometimes hard to determine why a logical layout changed, making
it hard to assign a meaningful name to the layout.

Processing
------------

See :ref:`developer` for background. We're going to need a "sheet process function."
This transforms the source sheet into the target collection, usually an output file.

The :py:func:`process_sheet` function is the heart of the application.
This handles all the rows present in a given sheet.

::

    def process_sheet( sheet, builder, persistence ):
        counts= defaultdict( int )
        if sheet.schema is None:
            # Empty sheet -- no embedded schema
            return counts
        for source_row in sheet.schema.rows_as_dict_iter(sheet):
            try:
                counts['read'] += 1
                row_dict= builder( source_row )
                f= ThisForm( **row_dict )
                if f.is_valid():
                    counts['processed'] += 1
                    this= f.save()
                    persistence.save_this( this )
                else:
                    counts['rejected'] += 1
            except Exception as e:
                counts['invalid'] += 1
                if persistence.stop_on_exception: raise
                summary= "{0} {1!r}".format( e.__class__.__name__, e.args )
                logger.error( summary )
                counts['error',summary] += 1
        return counts
        
Some applications will have variant processing for workbooks that
contain different types of sheets. 
This leads to different :py:func:`process_this_sheet` and :py:func:`process_that_sheet` functions. 
Each will follow the above template to process all rows of the sheet.

High-Level Interfaces
------------------------

These are the functions that can be used for live-file unit testing
of the application as a whole. The :py:func:`validate` function
uses a context manager for validation only. The :py:func:`process` function
uses the other context manager to that actual work is performed.

::

    def validate( sheet, builder ):
        with Validate_Only_Processing() as mode:
            counts= process_sheet( sheet, builder, mode )
        return counts

    def process( sheet, builder ):        
        with Persistent_Processing() as mode:
            counts= process_sheet( sheet, builder, mode )
        return counts

These higher-level functions share a common :py:func:`process_workbook`
function that does the real work.
     
::


    def process_workbook( source, sheet_func, builder_func ):
        for name in source.sheets():
            logger.info( "{0} :: {1}".format( input, name ) )
            sheet= source.sheet( name, 
                stingray.sheet.EmbeddedSchemaSheet,
                loader_class=stingray.schema.loader.HeadingRowSchemaLoader )
            counts= sheet_func( sheet, builder_func )
            logger.info( pprint.pformat(dict(counts)) )

When we do live-file testing of a given file, we can do something like the following.
This uses :py:func:`validate` to assure that the file's schema is correct.
See :ref:`demo_sqa` for more information.

..  parsed-literal::

    from some_app import validate
    class Test_Some_File( unittest.TestCase ):
        def setUp( self ):
            self.source= stingray.workbook.open_workbook( input )
            self.builder_func= builder_1
        def test_should_process_sheet1( self ):
            sheet= source.sheet( "Sheet1", 
                stingray.sheet.EmbeddedSchemaSheet,
                loader_class=stingray.schema.loader.HeadingRowSchemaLoader )
            counts= validate( sheet, self.builder_func )
            self.assertEqual( 12345, counts['read'] )            

Command-Line Interface
----------------------

We have some standard arguments.  
While we'd like to use "-v" for validate mode, this gets confused with setting the verbosity level.

::

    def parse_args():
        parser= argparse.ArgumentParser()
        parser.add_argument( 'file', nargs='+' )
        parser.add_argument( '-d', '--dry-run', default=False, action='store_true',  )
        parser.add_argument( '-l', '--layout', default='1', choices=('1','2') )
        parser.add_argument( '-v', '--verbose', dest='verbosity',
            default=logging.INFO, action='store_const', const=logging.DEBUG )
        return parser.parse_args()

The overall main program looks something like this.  It handles a number of
common main-program issues.

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
        builder_func= make_builder( args )
        sheet_func= validate if args.dry_run else process 
        logger.info( "Mode: {0}".format( sheet_func.__name__ ) )
        try:
            for input in args.file:
                with stingray.workbook.open_workbook( input ) as source:
                    process_workbook( source, sheet_func, builder_func )
            status= 0
        except Exception as e:
            logging.exception( e )
            status= 3 
        logging.shutdown()
        sys.exit( status )
        
Running the Demo
================

We can run this program like this.

..  code-block:: bash

    python3 demo/app.py -d -l 1 sample/\*.csv
    
This will apply builder with layout ``1`` against all of the :file:`sample/*.csv` files.

The output looks like this

..  code-block:: none

    INFO:demo/app.py:Mode: validate
    INFO:demo/app.py:sample/csv_workbook.csv :: csv_workbook
    INFO:demo/app.py:{'input': 2, 'valid': 2}
    INFO:demo/app.py:sample/simple.csv :: simple
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    ERROR:demo/app.py:KeyError 'Column "3" - string'
    INFO:demo/app.py:{'KeyError \'Column "3" - string\'': 7, 'input': 7}

We can see that :file:`sample/csv_workbook.csv` has two valid rows.

We can see that :file:`sample/simple.csv` has seven rows, all of which are missing the required value.  
If all the rows are wrong, the schema in the file
doesn't match the schema required by the application.
