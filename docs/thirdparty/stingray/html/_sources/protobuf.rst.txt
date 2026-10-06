..  #!/usr/bin/env python3

.. _`protobuf`:

###############################################################
Protobuf Module -- Unpacking iWork 13 files.
###############################################################

This is not a full implementation of Protobuf object representation.
This is a minimal implementation of protobuf parsing, enough to unpack iWork '13 files.

..  py:module:: protobuf 


The iWork '13 use of protobuf
===============================================

https://github.com/obriensp/iWorkFileFormat

https://github.com/obriensp/iWorkFileFormat/blob/master/Docs/index.md

    "Components are serialized into .iwa (iWork Archive) files, 
    a custom format consisting of a Protobuf stream wrapped in a Snappy stream.
   
    "Protobuf
   
    "The uncompresed IWA contains the Component's objects, serialized consecutively 
    in a Protobuf stream. Each object begins with a varint representing the length of 
    the ArchiveInfo message, followed by the ArchiveInfo message itself. 
    The ArchiveInfo includes a variable number of MessageInfo messages describing 
    the encoded Payloads that follow, though in practice iWork files seem to only 
    have one payload message per ArchiveInfo.
   
    "Payload
   
    "The format of the payload is determined by the type field of the associated 
    MessageInfo message. The iWork applications manually map these integer values 
    to their respective Protobuf message types, and the mappings vary slightly 
    between Keynote, Pages and Numbers. This information can be recovered by 
    inspecting the TSPRegistry class at runtime.

    "TSPRegistry"

    "The mapping between an object's MessageInfo.type and its respective Protobuf 
    message type must by extracted from the iWork applications at runtime. 
    Attaching to Keynote via a debugger and inspecting [TSPRegistry sharedRegistry] shows:

    "A full list of the type mappings can be found here."
   
    https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Persistence/MessageTypes

Message ``.proto`` files.

-   Table details 

    https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Messages/Proto/TSTArchives.proto
   
-   Numbers details

    https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Messages/Proto/TNArchives.proto
   
-   Calculating Engine details

    https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Messages/Proto/TSCEArchives.proto

-   Structure (i.e., TreeNode, perhaps more relevant for Keynote)

    https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Messages/Proto/TSKArchives.proto

We require two of the files from this project to map the internal code numbers 

-   https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Persistence/MessageTypes/Numbers.json

-   https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Persistence/MessageTypes/Common.json

These files are incorporated into this module as separate :file:`*.json` files.
See :ref:`installation` for more information on these files.

protobuf
===============================================
   
For more information on protobuf, see the following:

https://developers.google.com/protocol-buffers/

https://developers.google.com/protocol-buffers/docs/encoding

http://en.wikipedia.org/wiki/Protocol_Buffers

IWA Structure
================

Each IWA has an ArchiveInfo message.

..  parsed-literal::
   
    message ArchiveInfo {
        optional uint64 identifier = 1;
        repeated MessageInfo message_infos = 2;
    }

Within the ArchiveInfo is a MessageInfo message.

..  parsed-literal::
   
    message MessageInfo {
        required uint32 type = 1;
        repeated uint32 version = 2 [packed = true];
        required uint32 length = 3;
        repeated FieldInfo field_infos = 4;
        repeated uint64 object_references = 5 [packed = true];
        repeated uint64 data_references = 6 [packed = true];
    }

The MessageInfo is followed by the payload. That must be decoded to get the
actual data of interest.

Implementation
===============

Module docstring.

::

  """Read protobuf-serialized messages from IWA files used for Numbers '13 workbooks.

  https://developers.google.com/protocol-buffers/

  https://developers.google.com/protocol-buffers/docs/encoding

  Requires :file:`Numbers.json` and :file:`Common.json` from the installation
  directory.
  """

Some Overheads

::

  import logging
  import sys
  import os
  import json
  from collections import defaultdict, ChainMap
  from stingray.snappy import varint

..  py:class:: Message

    A definition of a generic protobuf message. This is both an instance
    and it also has staticmethods that build instances from a buffer of bytes.

    We don't use subclasses of ``Message``. The proper way to use
    Protobuf is to compile ``.proto`` files into Message class definitions.

::

  class Message:
      """Generic protobuf message built from sequence of bytes.
    
      :ivar name_: the protobuf message name.
      :ivar fields: a dict that maps field numbers to field values.
          The contained message objects are **not** parsed, but left as 
          raw bytes.
      """
      def __init__( self, name, bytes ):
          self.name_= name
          self.fields= Message.parse_protobuf(bytes)
      def __repr__( self ):
          return "{0}({1})".format( self.name_, self.fields )
      def __getitem__( self, index ):
          return self.fields.get(index,[])

..  py:method::  Message.parse_protobuf_iter( message_bytes )

    An iterative parser for the top-level (name, value) pairs in the protobuf stream.
    This yields all of the pairs that are parsed. This a static method which builds
    message instances.

::

      @staticmethod    
      def parse_protobuf_iter( message_bytes ):
          """Parse a protobuf stream, iterating over the name-value pairs that are present.
          This does NOT recursively descend through contained sub-messages.
          It does only the top-level message.
          """
          bytes_iter= iter(message_bytes)
          while True:
              try:
                  item= varint( bytes_iter )
              except StopIteration:
                  item= None
                  break
              field_number, wire_type = item >> 3, item & 0b111
              if wire_type == 0b000:   # varint representation
                  item_size= None # varint sizes vary, need something for debug message
                  field_value = varint( bytes_iter )
              elif wire_type == 0b001: # 64-bit == 8-byte
                  item_size= 8
                  field_value = tuple( next(bytes_iter) for i in range(8) )
              elif wire_type == 0b010: # varint length and then content
                  item_size= varint( bytes_iter )
                  field_value = tuple( next(bytes_iter) for i in range(item_size) )
              elif wire_type == 0b101: # 32-byte == 4-byte
                  item_size= 4
                  field_value = tuple( next(bytes_iter) for i in range(4) )
              else:
                  raise Exception( "Unsupported {0}: {1}, {2}", bin(item), field_number, wire_type )
              Message.log.debug(
                  '{0:b}, field {1}, type {2}, size {3}, = {4}'.format(
                  item, field_number, wire_type, item_size, field_value) )
              yield field_number, field_value
    
..  py:method::  Message.parse_protobuf( message_bytes )

    Create a bag in the form of a mapping ``{name: [value,value,value], ... }``. This will 
    contain the top-level identifiers and the bytes that could be used to parse 
    lower-level messages.

::

      @staticmethod    
      def parse_protobuf( message_bytes ):
          """Creates a bag of name-value pairs. Names can repeat, so values are 
          an ordered list.
          """
          bag= defaultdict( list )
          for name, value in Message.parse_protobuf_iter( message_bytes ):
              bag[name].append( value )
          return dict(bag)
    
A class-level logger. We don't want a logger for each instance, since we'll
create many ``Message`` instances.

::

  Message.log= logging.getLogger( Message.__class__.__qualname__ )

..  py:class:: Archive_Reader

    A Reader for IWA archives. This requires that the archive has been 
    processed by the :py:class:`snappy.Snappy` decompressor.

::

  class Archive_Reader:
      """Read and yield Archive entries from MessageInfo.type and payload.
      Resolves the ID into a protobuf message name.

      Mapping from types to messages

      -   https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Persistence/MessageTypes/Numbers.json

      -   https://github.com/obriensp/iWorkFileFormat/blob/master/iWorkFileInspector/iWorkFileInspector/Persistence/MessageTypes/Common.json
      """
      def __init__( self ):
          self.tsp_names= self._tsp_name_map()
          self.log= logging.getLogger( self.__class__.__qualname__ )


Mapping from internal code numbers of protobuf message class names.
This requires :file:`Numbers.json` and :file:`Common.json` from the installation
directory.


::

      @staticmethod
      def _tsp_name_map():
          """Build the TSPRegistry map from messageInfo.type to message proto
          """
          def load_map( filename ):
              """JSON documents have string keys: these must be converted to int."""
              installed= os.path.dirname(__file__)
              with open( os.path.join(installed, filename) ) as source:
                  raw= json.load( source )
              return dict( (int(key), value) for key, value in raw.items() )

          tsp_names= ChainMap(
              load_map("Numbers.json"),
              load_map("Common.json"),
          )     
          return tsp_names

..  py:method::  Archive_Reader.make_message( messageInfo, payload )

Create the payload message from a MessageInfo instance and the payload bytes.

::


      def make_message( self, messageInfo, payload ):
          name= self.tsp_names[messageInfo[1][0]]
          return Message( name, payload )
        
..  py:method::  Archive_Reader.archive_iter( data )

Iterate through all messages in this IWA archive. Locate the ArchiveInfo,
MessageInfo and Payload. Parse the payload to create the final message
that's associated with the ID in the ArchiveInfo.

::

      def archive_iter( self, data ):
          """Iterate through the iWork protobuf-serialized archive:
          Locate the ArchiveInfo object.
          Each ArchiveInfo contains a MessageInfo(s).
          Each MessageInfo describes a payload. 
          It appears that there's only one MessageInfo per ArchiveInfo
          even though the ``.proto`` file indicates multiple as possible.

          Yield a sequence of iWork archived messages as pairs: 
              identifier, Message object built from the payload.
          """
          protobuf= iter(data)
          while True:
              try:
                  size= varint( protobuf )
              except StopIteration:
                  size= None
                  break
              self.log.debug( "{0} bytes".format(size) )
              message_bytes= [ next(protobuf) for i in range(size) ]
              archiveInfo= Message( "ArchiveInfo", message_bytes)
              archiveInfo.identifier = archiveInfo[1][0]
              archiveInfo.message_infos = [ 
                  Message("MessageInfo", mi) for mi in archiveInfo[2] ]
              self.log.debug( " ArchiveInfo identifier={0} message_infos={1}".format(
                  archiveInfo.identifier, archiveInfo.message_infos) )
              messageInfo_0= archiveInfo.message_infos[0]
              messageInfo_0.length= messageInfo_0[3][0]
              self.log.debug( "   MessageInfo length={0}".format(messageInfo_0.length) )
              payload_raw= [ next(protobuf) for i in range(messageInfo_0.length) ]
              self.log.debug( "     Payload {0!r}".format(payload_raw) )
              message= self.make_message( messageInfo_0, payload_raw )
              yield archiveInfo.identifier, message
