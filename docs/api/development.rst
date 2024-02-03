:hidetoc: 1

Development
###########

Numbers File Formats
--------------------

Numbers uses a proprietary, compressed binary format to store its
tables. This format is comprised of a zip file containing images, as
well as `Snappy <https://github.com/google/snappy>`__-compressed
`Protobuf <https://github.com/protocolbuffers/protobuf>`__ ``.iwa``
files containing metadata, text, and all other definitions used in the
spreadsheet.

Protobuf updates
~~~~~~~~~~~~~~~~

As ``numbers-parser`` includes private Protobuf definitions extracted
from a copy of Numbers, new versions of Numbers will inevitably create
``.numbers`` files that cannot be read by ``numbers-parser``. As new
versions of Numbers are released, running ``make bootstrap`` will
perform all the steps necessary to recreate the protobuf files used
``numbers-parser`` to read Numbers spreadsheets.

The default protobuf package installation may not include the C++
optimized version which is required by the bootstrapping scripts to
extract protobufs. You will receive the following error during build if
this is the case:

``This script requires the Protobuf installation to use the C++ implementation. Please reinstall Protobuf with C++ support.``

To include the C++ support, download a released version of Google
protobuf `from github <https://github.com/protocolbuffers/protobuf>`__.
Build instructions are described in
```src/README.md`` <https://github.com/protocolbuffers/protobuf/blob/main/src/README>`__.These
have changed greatly over time, but as of April 2023, this was useful:

.. code:: shell

   bazel build :protoc :protobuf
   cmake . -DCMAKE_CXX_STANDARD=14
   cmake --build . --parallel 8
   export PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION=cpp
   export LD_LIBRARY_PATH=../bazel-bin/src/google
   cd python
   python3 setup.py -q bdist_wheel --cpp_implementation --warnings_as_errors --compile_static_extension

This can then be used ``make bootstrap`` in the ``numbers-parser``
source tree. The signing workflow assumes that you have an Apple
Developer Account and that you have created provisioning profile that
includes iCloud. Using a self-signed certificate does not seem to work,
at least on Apple Silicon (a working PR contradicting this is greatly
appreciated).

``make bootstrap`` requires
`PyObjC <https://pypi.org/project/pyobjc/>`__ to generate font maps, but
this dependency is excluded from Poetry to ensure that tests can run on
non-Mac OSes. You can run ``poetry run pip install PyObjC`` to get the
required packages.

Credits
-------

``numbers-parser`` was built by `Jon
Connell <http://github.com/masaccio>`__ but relies heavily on from
`prior work <https://github.com/psobot/keynote-parser>`__ by `Peter
Sobot <https://petersobot.com>`__ to read the IWA format archives used
by Apple’s iWork family of applications, and to regenerate the mapping
files required for Python. Both modules are derived from `previous
work <https://github.com/obriensp/iWorkFileFormat/blob/master/Docs/index.md>`__
by `Sean Patrick O’Brien <http://www.obriensp.com>`__.

Decoding the data structures inside Numbers files was helped greatly by
`Stingray-Reader <https://github.com/slott56/Stingray-Reader>`__ by
`Steven Lott <https://github.com/slott56>`__.

Formula tests were adapted from JavaScript tests used in
`fast-formula-parser <https://github.com/LesterLyu/fast-formula-parser>`__.

Decimal128 conversion to and from byte storage was adapted from work
done by the `SheetsJS project <https://github.com/SheetJS/sheetjs>`__.
SheetJS also helped greatly with some of the steps required to
successfully save a Numbers spreadsheet.