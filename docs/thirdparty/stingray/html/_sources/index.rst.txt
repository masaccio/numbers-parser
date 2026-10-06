.. The Stingray Schema-Based File Reader documentation master file, created by
   sphinx-quickstart on Sat Aug 27 11:46:24 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

================================================================
The Stingray Schema-Based File Reader
================================================================

The Stingray Reader tackles four fundamental issues in 
processing a file: 

-   How are the bytes organized?  What is the Physical Format?
-   Haw are the data objects organized?  What is the Logical Layout?
-   What do the bytes *mean*?  What is the Conceptual Content?
-   How can we assure ourselves that our applications will work with this file?

The problem we have is that the schema is not always bound
to a given file nor is the schema clearly bound to an application program.
There are two examples of this separation between schema and content:

-	We might have a spreadsheet where there aren't even column titles.

-	We might have a pure data file (for example from a legacy COBOL program)
	which is described by a separate schema.
	
One goal of good software is to cope reasonably well with variability
of user-supplied inputs.  Providing data by spreadsheet is 
often the most desirable choice for users.  In some cases, it's the
only acceptable choice.  Since spreadsheets are tweaked manually, they
may not have a simple, fixed schema or logical layout. 

A workbook (the container of individual sheets)
can be encoded in any of a number of physical
formats: XLS, CSV, XLSX, ODS to name a few.  We would like our applications
to be independent of these physical formats.  We'd like to focus
on the logical layout.

Data supplied in the form of a workbook can suffer from numerous data quality issues.  
We need to be assured that a file actually conforms to a required
schema.

A COBOL file parallels a workbook sheet in several ways. It also introduces
some unique complications. We'd like to provide a suite of tools that work
well with common spreadsheets as well as COBOL files, allowing some 
uniformity in processing various kinds of data.

Technology
==========

While this is 100% Python, it's not *simply* Python. The actual code
is built from this document. 

..  warning:: The Code Did Not Come First

    The document doesn't follow behind the code. The document **contains**
    the code. The code is extracted from the document. 
    For details, see the :ref:`build` section.

Contents
=========


.. toctree::
   :numbered:
   :maxdepth: 1
   
   introduction
   design
   package
   cell
   sheet
   schema
   schema_loader
   workbook/index
   iwork13
   cobol
   developer
   demo/index
   history
   testing/index
   build
   installation
   license

License
========

..	raw:: html

	<a rel="license" href="http://creativecommons.org/licenses/by-nc-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="http://i.creativecommons.org/l/by-nc-sa/4.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-nc-sa/4.0/">Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License</a>.

The TODO List
===============

..  todolist::

Indices and Tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

