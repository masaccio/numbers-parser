##################
Main Test Script
##################

This module imports the other test modules and builds a complete suite
from the individual module suites.

::

  """stingray test script."""
  import unittest
  import sys
  import logging
  import test.cell
  import test.sheet
  import test.schema
  import test.schema_loader
  import test.workbook
  import test.cobol
  import test.cobol_loader
  import test.cobol_2
  import test.snappy_protobuf
        
Construction of an overall suite depends on each module providing
and easy-to-use :py:func:`suite` function that returns the module's suite.

::

  def suite():
      s= unittest.TestSuite()
      s.addTests( test.cell.suite() )
      s.addTests( test.sheet.suite() )
      s.addTests( test.schema.suite() )
      s.addTests( test.schema_loader.suite() )
      s.addTests( test.workbook.suite() )
      s.addTests( test.cobol.suite() )
      s.addTests( test.cobol_loader.suite() )
      s.addTests( test.cobol_2.suite() )
      s.addTests( test.snappy_protobuf.suite() )
      return s
    
  def main():
      with test.Logger( stream=sys.stderr, level=logging.WARN ):
    
          # Single test case for debugging purposes
          #single= test.cobol_2.Test_Copybook_11
          #suite= unittest.defaultTestLoader.loadTestsFromTestCase(single)
          #return unittest.TextTestRunner().run(suite)
        
          # All test cases is the default
          return unittest.TextTestRunner().run(suite())

  if __name__ == "__main__":
      main()
