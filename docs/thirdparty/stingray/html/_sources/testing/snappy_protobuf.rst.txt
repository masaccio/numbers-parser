

.. _`test_snappy_protobuf`:

##########################
Test Snappy and Protobuf
##########################

These are tests for the snappy decompression and the protobuf object encoding.

Overheads
=============

::

  """stingray.snappy and stingray.protobuf Unit Tests."""
  import unittest
  import io
  import stingray.snappy
  import stingray.protobuf
  import logging, sys

Snappy Reader
==============

The snappy decompressor is built around several layers of protocol.

1. There's a framing protocol with a frame type and size.

2. There's a size for the uncompressed data in the frame.

3. There are the 4 kinds of snappy tags that create the data.

   - 00 : literal
   - 01, 10, 11 : copies with various kinds of sizes and offsets.

::

  class Test_Snappy( unittest.TestCase ):
      def setUp( self ):
          #logging.getLogger().setLevel( logging.DEBUG )
          self.buffer = io.BytesIO( 
              b'\x00\x12\x00\x00' # Header: 18 bytes in frame
              b'\x18' # Size of the uncompressed data
              b'\x14hi mom' # 000101,00 -> literal of 5+1=6 bytes
              b'\x09\x06' # 000,010,01 00000110 -> copy 4+2=6 bytes offset of 6
              b'\x16\x00\x0c' # 000101,10 00000000 00001100 -> copy 5+1=6 bytes offset of 12
              b'\x17\x00\x00\x00\x12' # 000101,11 00000000 00000000 00000000 00010010 -> copy 5+1=6 bytes offset of 18

              b'\x00\x0a\x00\x00' # Header: 10 bytes in frame
              b'\x08' # Size of the uncompressed data
              b'\x1CZYXW!@#$' # 000111,00 -> literal of 7+1=8 bytes
              )
          self.snappy= stingray.snappy.Snappy()
      def test_should_decompress( self ):
          data= self.snappy.decompress( self.buffer )
          self.assertEqual( b'hi momhi momhi momhi momZYXW!@#$', data )

Protobuf Decoder
=================

The protobuf decoder unpacks various kinds of data to create simple Message instances.

This is a two-layer protocol.

There's a protobuf-encoded ArchiveInfo message.

..  parsed-literal::
   
    message ArchiveInfo {
        optional uint64 identifier = 1;
        repeated MessageInfo message_infos = 2;
    }


This contains a protobuf-encoded MessageInfo message.

..  parsed-literal::

    message MessageInfo {
        required uint32 type = 1;
        repeated uint32 version = 2 [packed = true];
        required uint32 length = 3;
        repeated FieldInfo field_infos = 4;
        repeated uint64 object_references = 5 [packed = true];
        repeated uint64 data_references = 6 [packed = true];
    }

The message has a payload, which is the relevant message that is part of the Numbers
workbook.    

Here's the proto definition for the test message encoded in the payload below.

..  parsed-literal::

    message DocumentArchive {
      repeated .TSP.Reference sheets = 1;
      required .TSA.DocumentArchive super = 8;
      optional .TSP.Reference calculation_engine = 3 [deprecated = true];
      required .TSP.Reference stylesheet = 4;
      required .TSP.Reference sidebar_order = 5;
      required .TSP.Reference theme = 6;
      optional .TN.UIStateArchive uistate = 7;
      optional .TSP.Reference custom_format_list = 9;
      optional string printer_id = 10;
      optional string paper_id = 11;
      optional .TSP.Size page_size = 12;
    }
   
Note that each part of this (except for ``printer_id`` and ``paper_id``) 
could involve decoding a contained message. We don't recursively descend, however,
merely decoding the top message. 


::

  class Test_Protobuf( unittest.TestCase ):
      def setUp( self ):
          #logging.getLogger("Archive_Reader").setLevel( logging.DEBUG )
          self.buffer= bytes( (
              37, # 37 bytes follow.
              # AchiveInfo and MessageInfo
              8, 1, 18, 33, 8, 1, 18, 3, 1, 0, 5, 24, 91, 42, 22, 133, 12, 170, 15, 134, 
              12, 185, 14, 132, 12, 137, 12, 135, 12, 183, 14, 136, 12, 184, 15, 183, 15,
              # Payload
              10, 3, 8, 183, 14, 10, 3, 8, 170, 15, 10, 3, 8, 183, 15, 34, 3, 8, 136, 12, 
              42, 3, 8, 185, 14, 50, 3, 8, 137, 12, 66, 33, 10, 5, 58, 3, 8, 132, 12, 26, 
              2, 101, 110, 34, 3, 8, 133, 12, 42, 3, 8, 184, 15, 50, 3, 8, 135, 12, 58, 
              3, 8, 134, 12, 64, 0, 82, 1, 32, 90, 9, 110, 97, 45, 108, 101, 116, 116, 
              101, 114, 98, 10, 13, 0, 0, 25, 68, 21, 0, 0, 70, 68,
              )
          )
          self.reader= stingray.protobuf.Archive_Reader()
      def test_should_decode( self ):
          message_list = list( self.reader.archive_iter(self.buffer) )
          m0_id, m0_m = message_list[0]
          self.assertEqual( 1, m0_id )
          self.assertEqual( "TN.DocumentArchive", m0_m.name_ )
          sheets= m0_m[1]
          self.assertEqual( [(8, 183, 14), (8, 170, 15), (8, 183, 15)], sheets )
          super= m0_m[8]
          self.assertEqual( [(10, 5, 58, 3, 8, 132, 12, 26, 2, 101, 110, 34, 3, 8, 133, 12, 42, 3, 8, 184, 15, 50, 3, 8, 135, 12, 58, 3, 8, 134, 12, 64, 0)], super )
          calculation_engine= m0_m[3]
          self.assertEqual( [], calculation_engine )
          stylesheet= m0_m[4]
          self.assertEqual( [(8, 136, 12)], stylesheet )
          sidebar_order= m0_m[5]
          self.assertEqual( [(8, 185, 14)], sidebar_order )
          theme= m0_m[6]
          self.assertEqual( [(8, 137, 12)], theme )
          uistate= m0_m[7]
          self.assertEqual( [], uistate )
          custom_format_list= m0_m[9]
          self.assertEqual( [], custom_format_list )
          printer_id= m0_m[10]
          self.assertEqual( [(32,)], printer_id )
          paper_id= m0_m[11]
          # b'na-letter' is the paper_id
          self.assertEqual( [(110, 97, 45, 108, 101, 116, 116, 101, 114)], paper_id )
          page_size= m0_m[12]
          self.assertEqual( [(13, 0, 0, 25, 68, 21, 0, 0, 70, 68)], page_size )

Test Suite and Runner
=====================

In case we want to build up a larger test suite, we avoid doing
any real work unless this is the main module being executed.

::

  import test
  suite= test.suite_maker( globals() )

  if __name__ == "__main__":
      print( __file__ )
      unittest.TextTestRunner().run(suite())
