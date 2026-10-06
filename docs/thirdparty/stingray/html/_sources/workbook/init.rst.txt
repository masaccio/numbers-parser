..  #!/usr/bin/env python3

.. _`workbook_init`:

###############################################################
Workbook __init__ Module -- Wrapper for all implementations
###############################################################
   
A few Python overheads that we put in the ``__init__``
module of this package. Our goal is to make it so that only
the top-level package is imported; the individual workbook modules
are not generally expected to be used by an application.

..  py:module:: workbook

::

  """stingray.workbook -- Opens workbooks in various
  formats, binds their associated schema, accesses them as Sheets with
  Rows and Cells.

  This is a kind of **Wrapper** or **Facade** that unifies :py:mod:`csv` and
  :py:mod:`xlrd`. It handles a number of file formats including
  :file:`.xlsx`, :file:`.ods`, and Numbers.
  """

In order top open files of various types, we'll bring in a number of
helpful modules.

::

  import xml.etree.cElementTree as dom
  from collections import defaultdict
  import zipfile
  import datetime
  from io import open
  import os.path
  import pprint
  import re
  import glob
  import logging
  import decimal

We'll rely on definitions of :py:mod:`cell`, :py:mod:`sheet`,
and :py:mod:`schema.loader`. We have an implicit dependency on :py:mod:`schema`: 
we'll be given schema objects to work with.

::

  import stingray.cell
  import stingray.sheet
  import stingray.schema.loader

We'll explicitly import the top-level class definition for each
flavor of Workbook we can support. Because we import these here, these
classes will be available with a simple import of :mod:`workbook`.

::

  from stingray.workbook.csv import CSV_Workbook
  from stingray.workbook.xlsx import XLSX_Workbook
  from stingray.workbook.ods import ODS_Workbook
  from stingray.workbook.numbers_09 import Numbers09_Workbook
  from stingray.workbook.numbers_13 import Numbers13_Workbook
  from stingray.workbook.fixed import Fixed_Workbook

Exceptions
==========


..  py:class:: UnknownFormat

    The :py:class:`UnknownFormat` exception is raised when a workbook can't be
    opened.
   
::

  class UnknownFormat( Exception ):
      """The workbook can't be opened."""
      pass

..  py:class:: No_Schema

    The :py:class:`No_Schema` exception is raised if there's a problem
    locating an external schema for a workbook.

::

  class No_Schema( Exception ):
      """A valid schema could not be loaded."""
      pass

Optional Modules
=================

The :py:mod:`workbook.package.xls` module depends on :py:mod:`xlrd`.
https://pypi.python.org/pypi/xlrd/0.9.2 http://www.lexicon.net/sjmachin/xlrd.htm

We can't guarantee that :py:mod:`xlrd` is available. Also, old :file:`.xls` files are 
becoming less frequently used, so we're making this optional.

::

  try:
      from stingray.workbook.xls import XLS_Workbook
  except ImportError:
      from stingray.workbook.base import Workbook
      class XLS_Workbook( Workbook ):
          """No ``xlrd`` Available."""
          def __init__( self, *args, **kw ):
              raise UnknownFormat


Workbook Subclasses
=====================

We have a number of concrete subclasses of :py:class:`workbook.base.Workbook`. 
These are imported from submodules and made visible in this module.

-   :py:class:`workbook.CSV_Workbook`.  This is a degenerate case, where the workbook appears to contain
    a single sheet.  This sheet is the CSV file, accessed via the built-in
    :py:func:`csv.reader`.

-   :py:class:`workbook.XLS_Workbook`.  This is the workbook as processed by :py:mod:`xlrd`.  These classes
    wrap :py:mod:`xlrd` classes to which the real work is delegated.
    This is optional -- if :py:mod:`xlrd` is not installed, things will work,
    but these files cannot be opened.

-   :py:class:`workbook.XLSX_Workbook`.  This is the workbook after unzipping and using an XML parser
    on the various document parts.  Mostly, this is a matter of unzipping
    and parsing parts of the document to create a DOM which can be traversed
    as needed.

-   :py:class:`workbook.Numbers09_Workbook`.
    This handles the iWork '09 Numbers files with multiple 
    workspaces and multiple tables in each workspace. 

-   :py:class:`workbook.Numbers13_Workbook`
    These handle the iWork '13 Numbers files with multiple 
    workspaces and multiple tables in each workspace. 
   
-   :py:class:`workbook.ODS_Workbook`.

-   :py:class:`workbook.Fixed_Workbook`.  This is actually a fairly complex case.  The workbook will appear to
    contain a single sheet; this sheet is the fixed format file.  Schema information
    was required up front, unlike the other formats.

Further extensions will handle various kinds of COBOL files. They're similar to Fixed Workbooks.
See :ref:`cobol`.

Each of these is a context manager, so we include the necessary methods.

Note that workbooks are rarely simple files.  Sometimes they are ZIP archive
members.  Sometimes, they must be processed via **gzip**. Sometimes they involve
Snappy compression.

In order to minimize the assumptions, we try to handle two forms of file processing:

-   By name. In this case, the file name is provided. The file is opened and closed by
    the Workbook using the context manager interface.

-   By file-like object. An open file-like object is provided. No additional
    context management is performed. This is appropriate when a workbook is itself
    a member of a larger archive.


Workbook Factory
=================

This is the factory which creates a subclass of :class:`Workbook` for a 
a given file. 

..  py:class:: Opener

    An opener **Factory** class.  A subclass can extend this to handle other file
    extensions and physical formats.

::

  class Opener:
      """An extensible opener that examines the file extension and locates
      a proper Workbook subclass.
      """
      def __call__( self, name, file_object=None,
          schema_path='.', schema_sheet= None, **kw ):
          """Open a workbook.
        
          :param name: filename to open.
          :param file_object: File-like object to process.  If not
          provided the named file will be opened.
          :keyword schema_path: Directory with external schema files
          :keyword schema_sheet: A sheet in an external schema workbook.
          """
          _, ext = os.path.splitext( name )
          ext = ext.lower()
          if ext == ".xls": 
              return XLS_Workbook( name, file_object )
          elif ext in ( ".xlsx", ".xlsm" ):
              return XLSX_Workbook( name, file_object )
          elif ext in ( ".csv", ):
              return CSV_Workbook( name, file_object, **kw )
          elif ext in ( ".tab", ):
              return CSV_Workbook( name, file_object, delimiter='\t', **kw )
          elif ext in ( ".ods", ):
              return ODS_Workbook( name, file_object )
          elif ext in ( ".numbers", ):
              # Directory? It's Numbers13_Workbook; Zipfile? It's Numbers09_Workbook
              if os.path.is_dir( name ):
                  return Numbers13_Workbook( name, file_object )
              else:
                  return Numbers09_Workbook( name, file_object )
          else:
              # Fixed format files with no specific extension
              # Ideally :file:`somefile.schema` is the file
              # and :file:`schema.csv` or :file:`schema.xlsx` can be tracked down.
              schema_pat= os.path.join(schema_path, ext[1:]+".*")
              schema_choices= glob.glob( schema_pat )
              if schema_choices:
                  schema_name= schema_choices[0]
                  schema_wb= open_workbook( schema_name )
                  esl= stingray.schema.loader.ExternalSchemaLoader( schema_wb, schema_sheet )
                  schema= esl.schema()
                  return Fixed_Workbook( name, file_object, schema=schema )
              else:
                  raise No_Schema( schema_pat )

..  py:function:: open_workbook( name, file_object, schema_path, schema_sheet )

    Open a workbook.
   
    For fixed format files, we attempt to track down and load the relevant
    schema file. The idea here is that a file's extension can map to the schema's
    filename.
   
    :samp:`somefile.{schema}` would use a :samp:`{schema}.csv` workbook as it's schema.
    We'll simply try the first file that matches :samp:`{schema}.*` to see if it's
    a workbook we can open.
   
    :param name: The name of the file.
   
    :param file_object: (optional) already opened file object.
   
    :param schema_path: (optional) filename for an external schema file.
   
    :param schema_sheet: (optional) name of a sheet with a schema.
   
For fixed format files, we attempt to track down and load the relevant
schema.  An application might have narrower and more specific rules
for binding file and schema.  

The default :py:func:`workbook.open_workbook` is simply an instance
of the :py:class:`workbook.Opener`.

::

  open_workbook= Opener()

When creating a subclass, use the **Chain of Command** pattern.
This allows a user to create subclasses to handle the various other file name extensions.
Here's an example:

..  parsed-literal::

    class MyOpener(workbook.Opener):
        def __call__(self, name, file_object=None,
            schema_path='.', schema_sheet=None, \*\*kw ):
            if fnmatch(name, "\*.dat"):
                esl= stingray.schema.loader.ExternalSchemaLoader( 
                    os.path.join(schema_path, "schemafile.csv") )
                schema= esl.schema()
                return CSV_Workbook( name, file_object, schema=schema, delimiter="|" )
            return super().__call__(name, file_object, schema_path, schema_sheet, \*\*kw )

There may be application-specific rules, or command-line options that
will determine a mapping bewtween filename and physical format or filename and schema.



