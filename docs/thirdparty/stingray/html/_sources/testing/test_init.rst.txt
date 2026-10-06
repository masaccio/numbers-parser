..  #!/usr/bin/env python3

#############################
stingray.test ``__init__.py``
#############################

Overheads.

::

  """Stingray Schema-based File Reader Testing Package"""

  import unittest
  import logging, sys

This is a slightly modified version of the test discovery function
in the Python :mod:`unittest` package. 

::

    
  def suite_maker( source ):
      def make_suite():
          tests= unittest.TestSuite()
          for name in source:
              obj = source[name]
              if isinstance(obj, type) and issubclass(obj, unittest.TestCase):
                  tests.addTests(
                      unittest.defaultTestLoader.loadTestsFromTestCase(obj))
          return tests
      return make_suite
    
A handy logging context.

::

  class Logger:
      def __init__( self, **kw ):
          self.args= kw
      def __enter__( self ):
          logging.basicConfig( **self.args )
      def __exit__( self, *exc ):
          logging.shutdown()

We can use ``with test.Logger( stream=sys.stderr, level=logging.DEBUG ):``
To enable logging.