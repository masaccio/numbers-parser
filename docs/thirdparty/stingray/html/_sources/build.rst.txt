..  #!/usr/bin/env python3

..  _`build`:

#########################
Stingray Build
#########################

The source for :py:mod:`stingray` is a Sphinx project that depends on PyLit.
Yes.  The documentation spawns the code.

..  note:: PyLit Feature

    The first line of each file must be ``..    #!/usr/bin/env python3``.
  
    The four spaces between ``..`` and ``#!`` defines the indent used for
    each line of code in the file.  
  
    Four spaces.

Stingray depends on the following

-   xlrd.  http://www.lexicon.net/sjmachin/xlrd.htm

    Version 0.9.2 is Python 3 compatible. https://pypi.python.org/pypi/xlrd/0.9.2

In addition to Python 3.3, there are two other projects used to build.

-   PyLit.  https://github.com/slott56/PyLit-3

-   Sphinx.  http://sphinx.pocoo.org/

The PyLit install is little more than a download and move the :file:`pylit.py` file to
the Python :file:`site-packages` directory.

..  code-block:: bash

    python3 -m pip install xlrd sphinx

The diagrams are done with YUML. See http://yuml.me.

Build Procedure
==================

1.  Bootstrap the :file:`build.py` script by running PyLit.

    ..  code-block:: bash
  
        python3 -m pylit -t source/build.rst build.py

    This reports that an extract was written to :file:`build.py`.
  
2.  Use the :file:`build.py` script to create the ``stingray`` source, unit
    tests, demonstration applications.  
    Build the Sphinx documentation.  
    And run the unit test, too.

    ..  code-block:: bash
  
        python3 build.py --makepy
      
    At the end of this step, the directory tree will include the following.
  
    -   :file:`build`.  The documentation.  In HTML.
    -   :file:`demo`.   Some demonstration applications that use ``stingray``. 
        See :ref:`demo`.
    -   :file:`stingray`.  The Python library, ready for installation.
    -   :file:`test`.  The unit test script.

    This reports, also, that 186 tests were run.
  

  
Build Script Design
=====================

This is a platform-independent :file:`build.py` file for the build script.
This can use ``from sphinx.application import Sphinx``
and ``import pylit`` to access these modules from within Python
instead of using command-line shell script techniques.

Overheads
-------------

We're going to make use of some "applications".

-   Sphinx top-level application.

-   PyLit top-level application.


::

  """Platform-independent build script"""
  from pathlib import Path
  import sys
  from sphinx.application import Sphinx
  import pylit
  import unittest
  import logging
  import shutil
  import argparse

Sphinx Build
---------------

..  py:function:: sphinx_build( srcdir, outdir, buildername='html' )

    Handle the simple use case for the ``sphinx-build`` script.

::

  def sphinx_build( srcdir, outdir, buildername='html' ):
      """Essentially: ``sphinx-build $* -b html source docs/html``"""
      confdir= srcdir= Path( srcdir ).resolve()
      outdir= Path( outdir ).resolve()
      doctreedir = outdir / '.doctrees'
      app = Sphinx(srcdir, confdir, outdir, doctreedir, buildername)
      app.build(force_all=False, filenames=[])
      return app.statuscode

PyLit Build Python
------------------

..  py:function:: pylit_build_py( srcdir, outdir )

    Handle the simple use case for PyLit transformation from RST to Python.

::

  def pylit_build_py( pyFile, rstFile ):
      """Essentially: ``python3 -m pylit -t source/demo/data_quality.rst demo/test.py``
      """
      pylit.main( txt2code= True, overwrite="update", infile= pyFile, outfile= pylit, args=[] )

PyLit Build RST
------------------

..  py:function:: pylit_build_rst( srcdir, outdir )

    Handle the other simple use case for PyLit transformation from Python to RST.

::

  def pylit_build_rst( pyFile, rstFile ):
      """Essentially: ``python3 -m pylit -c demo/test.py source/demo/data_quality.rst``
      """
      try:
          pylit.main( txt2code= False, overwrite="yes", infile= rstFile, outfile= pyFile, args=[] )
      except IOError as ex:
          print("Did not overwrite {pyFile}")


Copy Data File
---------------

..  py:function:: copy_file( srcdir, outdir )

    Handles the simple use case for copying a file.

::

  def copy_file( srcdir, outdir ):
      """Essentially: ``cp $srcdir $outdir``"""
      shutil.copy2( srcdir, outdir )

Copy Data File Back
-------------------

..  py:function:: copy_file_back( srcdir, outdir )

    Handles the reverse copy use case. This refuses to overwrite a file, unless it's newer.

::

  def copy_file_back(outdir, srcdir):
      """Essentially: ``if [$outdir -nt $srcsrc] then; cp $outdir $srcdir; fi``"""
      raise NotImplementedError


Run the Test Script
-----------------------

..  py:function:: run_test( )

    In effect, this does ``python3 test/main.py``

::

  def run_test():
      import test.main 
      result= test.main.main()
      if result.failures:
          sys.exit(result.failures)

The "Forward" Build Sequence from RST to PY
-------------------------------------------

Because of some file renames and slight changes in structure, this isn't
a trivial walk through a directory tree.

::
    
  def build_py():
      (Path.cwd()/"stingray"/"schema").mkdir(exist_ok=True)
      (Path.cwd()/"stingray"/"cobol").mkdir(exist_ok=True)
      (Path.cwd() / "stingray" / "workbook").mkdir(exist_ok=True)

      pylit_build_py( 'source/stingray_init.rst', 'stingray/__init__.py' )
      pylit_build_py( 'source/cell.rst', 'stingray/cell.py' )
      pylit_build_py( 'source/sheet.rst', 'stingray/sheet.py' )
      pylit_build_py( 'source/workbook/init.rst', 'stingray/workbook/__init__.py' )
      pylit_build_py( 'source/workbook/base.rst', 'stingray/workbook/base.py' )
      pylit_build_py( 'source/workbook/csv.rst', 'stingray/workbook/csv.py' )
      pylit_build_py( 'source/workbook/xls.rst', 'stingray/workbook/xls.py' )
      pylit_build_py( 'source/workbook/xlsx.rst', 'stingray/workbook/xlsx.py' )
      pylit_build_py( 'source/workbook/ods.rst', 'stingray/workbook/ods.py' )
      pylit_build_py( 'source/workbook/numbers_09.rst', 'stingray/workbook/numbers_09.py' )
      pylit_build_py( 'source/workbook/numbers_13.rst', 'stingray/workbook/numbers_13.py' )
      pylit_build_py( 'source/workbook/fixed.rst', 'stingray/workbook/fixed.py' )
      pylit_build_py( 'source/schema.rst', 'stingray/schema/__init__.py' )
      pylit_build_py( 'source/schema_loader.rst', 'stingray/schema/loader.py' )
      pylit_build_py( 'source/cobol_init.rst', 'stingray/cobol/__init__.py' )
      pylit_build_py( 'source/cobol_loader.rst', 'stingray/cobol/loader.py' )
      pylit_build_py( 'source/cobol_defs.rst', 'stingray/cobol/defs.py' )
      pylit_build_py( 'source/snappy.rst', 'stingray/snappy.py' )
      pylit_build_py( 'source/protobuf.rst', 'stingray/protobuf.py' )
      pylit_build_py( 'source/installation.rst', 'setup.py' )
  
      copy_file( 'source/Numbers.json', 'stingray/Numbers.json' )
      copy_file( 'source/Common.json', 'stingray/Common.json' )

      (Path.cwd()/"test").mkdir(exist_ok=True)

      pylit_build_py( 'source/testing/test_init.rst', 'test/__init__.py' )
      pylit_build_py( 'source/testing/main.rst', 'test/main.py' )
      pylit_build_py( 'source/testing/cell.rst', 'test/cell.py' )
      pylit_build_py( 'source/testing/sheet.rst', 'test/sheet.py' )
      pylit_build_py( 'source/testing/schema.rst', 'test/schema.py' )
      pylit_build_py( 'source/testing/schema_loader.rst', 'test/schema_loader.py' )
      pylit_build_py( 'source/testing/workbook.rst', 'test/workbook.py' )
      pylit_build_py( 'source/testing/cobol.rst', 'test/cobol.py' )
      pylit_build_py( 'source/testing/cobol_loader.rst', 'test/cobol_loader.py' )
      pylit_build_py( 'source/testing/cobol_2.rst', 'test/cobol_2.py' )
      pylit_build_py( 'source/testing/snappy_protobuf.rst', 'test/snappy_protobuf.py' )
  
      (Path.cwd()/"demo").mkdir(exist_ok=True)

      pylit_build_py( 'source/demo/data_quality.rst', 'demo/test.py' )
      pylit_build_py( 'source/demo/validation.rst', 'demo/app.py' )
      pylit_build_py( 'source/demo/profile.rst', 'demo/profile.py' )
      pylit_build_py( 'source/demo/cobol_reader.rst', 'demo/cobol_reader.py' )
  
      run_test()

      sphinx_build( 'source', 'docs/html', 'html' )

The "Reverse" Build Sequence from PY to RST
-------------------------------------------

::

  def build_rst():
      (Path.cwd()/"source"/"demo").mkdir(exist_ok=True)
      (Path.cwd()/"source"/"testing").mkdir(exist_ok=True)
      (Path.cwd() / "source" / "workbook").mkdir(exist_ok=True)

      print("Unit test...")
      run_test()

      print("Building source...")
      pylit_build_rst('source/installation.rst', 'setup.py')
      pylit_build_rst('source/build.rst', 'build.py')

      pylit_build_rst('source/stingray_init.rst', 'stingray/__init__.py')
      pylit_build_rst('source/cell.rst', 'stingray/cell.py')
      pylit_build_rst('source/sheet.rst', 'stingray/sheet.py')
      pylit_build_rst('source/workbook/init.rst', 'stingray/workbook/__init__.py')
      pylit_build_rst('source/workbook/base.rst', 'stingray/workbook/base.py')
      pylit_build_rst('source/workbook/csv.rst', 'stingray/workbook/csv.py')
      pylit_build_rst('source/workbook/xls.rst', 'stingray/workbook/xls.py')
      pylit_build_rst('source/workbook/xlsx.rst', 'stingray/workbook/xlsx.py')
      pylit_build_rst('source/workbook/ods.rst', 'stingray/workbook/ods.py')
      pylit_build_rst('source/workbook/numbers_09.rst', 'stingray/workbook/numbers_09.py')
      pylit_build_rst('source/workbook/numbers_13.rst', 'stingray/workbook/numbers_13.py')
      pylit_build_rst('source/workbook/fixed.rst', 'stingray/workbook/fixed.py')
      pylit_build_rst('source/schema.rst', 'stingray/schema/__init__.py')
      pylit_build_rst('source/schema_loader.rst', 'stingray/schema/loader.py')
      pylit_build_rst('source/cobol_init.rst', 'stingray/cobol/__init__.py')
      pylit_build_rst('source/cobol_loader.rst', 'stingray/cobol/loader.py')
      pylit_build_rst('source/cobol_defs.rst', 'stingray/cobol/defs.py')
      pylit_build_rst('source/snappy.rst', 'stingray/snappy.py')
      pylit_build_rst('source/protobuf.rst', 'stingray/protobuf.py')

      # TODO: special case -- only copy the file back to the source if it's newer.
      # copy_file_back('source/Numbers.json', 'stingray/Numbers.json')
      # copy_file_back('source/Common.json', 'stingray/Common.json')

      (Path.cwd() / "test").mkdir(exist_ok=True)

      pylit_build_rst('source/testing/test_init.rst', 'test/__init__.py')
      pylit_build_rst('source/testing/main.rst', 'test/main.py')
      pylit_build_rst('source/testing/cell.rst', 'test/cell.py')
      pylit_build_rst('source/testing/sheet.rst', 'test/sheet.py')
      pylit_build_rst('source/testing/schema.rst', 'test/schema.py')
      pylit_build_rst('source/testing/schema_loader.rst', 'test/schema_loader.py')
      pylit_build_rst('source/testing/workbook.rst', 'test/workbook.py')
      pylit_build_rst('source/testing/cobol.rst', 'test/cobol.py')
      pylit_build_rst('source/testing/cobol_loader.rst', 'test/cobol_loader.py')
      pylit_build_rst('source/testing/cobol_2.rst', 'test/cobol_2.py')
      pylit_build_rst('source/testing/snappy_protobuf.rst', 'test/snappy_protobuf.py')

      print("Building html...")
      sphinx_build( 'source', 'docs/html', 'html' )

Which Direction?
----------------

There are two potential directions, depending on where the author was working.
We don't try to resolve individual files to allow mixing and matching of changes.

To do that, we'll parse one command-line argument to get the direction.

::

  def get_options(argv = sys.argv[1:]):
      parser = argparse.ArgumentParser()
      parser.add_argument("--makepy", "-p",
                          action='store_true',
                          help="Makes Python code from RST source")
      parser.add_argument("--makesrc", "-s",
                          action='store_true',
                          help="Rebuilds RST source after Python changes")
      options = parser.parse_args(argv)
      if options.makepy and not options.makesrc:
          return options
      elif options.makesrc and not options.makepy:
          return options
      parser.error("Choose one of --makepy or --makesrc")
      # Now. Clear out the values so they don't confuse pylit3
      sys.argv = sys.argv[:1]

Main Program Switch
---------------------

When the :file:`build.py` script is run from the command line,
it will execute the :py:func:`build` function.  When it is imported,
however, it will do nothing special.

::

  if __name__ == "__main__":
      options = get_options()
      if options.makepy:
          build_py()
      elif options.makesrc:
          build_rst()
      else:
          sys.exit("No valid option chosen.")
  
Additional Builds
=====================

Sometimes it's desriable to refresh the documentation.

The HTML pages are built with this command.

..  code-block:: bash

    sphinx-build -b html source docs/html
  
The LaTeX document is built with this command.

..  code-block:: bash

    sphinx-build -b latex source docs/latex
