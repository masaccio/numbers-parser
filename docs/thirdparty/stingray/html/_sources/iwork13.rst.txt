.. _`other_modules`:

#########################################
The "Other" Modules: snappy and protobuf
#########################################

There are two digressive modules that are part of reading iWork '13 Numbers files.
We could depend on other implementations. Instead, we've provided our own
implementation to reduce the dependencies.

It's sensible to use more sophisticated imports in the workbook package to
properly handle various kinds of Snappy compression and Protobuf representation 
implementations.
In the same way that the various XML workbooks could be built on any of
the Python XML modules.

However, it's somewhat simpler for developers to rely only on 
Stingray and (optionally) :mod:`xlrd`.

..  toctree::
    :maxdepth: 1

    snappy
    protobuf
