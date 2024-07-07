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
`pre-compiled binary libraries <https://github.com/cgohlke/win_arm64-wheels/>`__ which are only available
for Windows on Arm. There appear to be no x86 pre-compiled packages for Windows.

.. code:: text

   pip install python_snappy-0.6.1-cp312-cp312-win_arm64.whl