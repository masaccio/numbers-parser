Installation
------------

.. code:: bash

   python3 -m pip install numbers-parser

A pre-requisite for this package is `python-snappy <https://pypi.org/project/python-snappy/>`__ 
which will be installed by Python automatically, but python-snappy also requires that the binary
libraries for snappy compression are present.

The most straightforward way to install the binary dependencies is to use
`Homebrew <https://brew.sh>`__ and source Python from Homebrew rather than from macOS as described 
in the `python-snappy github <https://github.com/andrix/python-snappy>`__:

For Intel Macs:

.. code:: bash

   brew install snappy python3
   CPPFLAGS="-I/usr/local/include -L/usr/local/lib" \
   python3 -m pip install python-snappy

For Apple Silicon Macs:

.. code:: bash

   brew install snappy python3
   CPPFLAGS="-I/opt/homebrew/include -L/opt/homebrew/lib" \
   python3 -m pip install python-snappy

For Linux (your package manager may be different):

.. code:: bash

   sudo apt-get -y install libsnappy-dev

On Windows, you will need to either arrange for snappy to be found for VSC++ or you can install python
binary libraries compiled by `Christoph Gohlke 
<https://www.lfd.uci.edu/~gohlke/pythonlibs/#python-snappy>`__. You must select the correct python
version for your installation. For example for python 3.11:

.. code:: text

   pip install python_snappy-0.6.1-cp311-cp311-win_amd64.whl