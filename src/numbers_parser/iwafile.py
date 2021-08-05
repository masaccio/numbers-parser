# Forked from https://github.com/psobot/keynote-parser/blob/master/keynote_parser/codec.py

import struct
import snappy
from functools import partial

from numbers_parser.mapping import ID_NAME_MAP

from google.protobuf.internal.decoder import _DecodeVarint32
from google.protobuf.json_format import MessageToDict

from numbers_parser.generated.TSPArchiveMessages_pb2 import ArchiveInfo


class IWAFile(object):
    def __init__(self, chunks, filename=None):
        self.chunks = chunks
        self.filename = filename

    @classmethod
    def from_buffer(cls, data, filename=None):
        try:
            chunks = []
            while data:
                chunk, data = IWACompressedChunk.from_buffer(data, filename)
                chunks.append(chunk)

            return cls(chunks, filename)
        except Exception as e:  # pragma: no cover
            if filename:
                raise ValueError("Failed to deserialize " + filename) from e
            else:
                raise

    def to_dict(self):
        try:
            return {"chunks": [chunk.to_dict() for chunk in self.chunks]}
        except Exception as e:  # pragma: no cover
            if self.filename:
                raise ValueError("Failed to serialize " + self.filename) from e
            else:
                raise


class IWACompressedChunk(object):
    def __init__(self, archives):
        self.archives = archives

    def __eq__(self, other):
        return self.archives == other.archives  # pragma: no cover

    @classmethod
    def _decompress_all(cls, data):
        while data:
            header = data[:4]

            first_byte = header[0]
            if not isinstance(first_byte, int):
                first_byte = ord(first_byte)

            if first_byte != 0x00:
                raise ValueError(  # pragma: no cover
                    "IWA chunk does not start with 0x00! (found %x)" % first_byte
                )

            unpacked = struct.unpack_from("<I", bytes(header[1:]) + b"\x00")
            length = unpacked[0]
            chunk = data[4 : 4 + length]
            data = data[4 + length :]

            try:
                yield snappy.uncompress(chunk)
            except Exception:  # pragma: no cover
                # Try to see if this data isn't compressed in the first place.
                # If this data is still compressed, parsing it as Protobuf
                # will almost definitely fail anyways.
                yield chunk

    @classmethod
    def from_buffer(cls, data, filename=None):
        data = b"".join(cls._decompress_all(data))
        archives = []
        while data:
            archive, data = IWAArchiveSegment.from_buffer(data, filename)
            archives.append(archive)
        return cls(archives), None

    def to_dict(self):
        return {"archives": [archive.to_dict() for archive in self.archives]}


class ProtobufPatch(object):
    def __init__(self, data):
        self.data = data

    def __eq__(self, other):
        return self.data == other.data  # pragma: no cover

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.data)  # pragma: no cover

    def to_dict(self):
        return message_to_dict(self.data)

    @classmethod
    def FromString(cls, message_info, proto_klass, data):
        # Recent numbers does not apear to store date this way
        assert len(message_info.diff_field_path.path) != 1
        return cls(proto_klass.FromString(data))


class IWAArchiveSegment(object):
    def __init__(self, header, objects):
        self.header = header
        self.objects = objects

    def __eq__(self, other):
        return (
            self.header == other.header and self.objects == other.objects
        )  # pragma: no cover

    def __repr__(self):
        return "<%s identifier=%s objects=%s>" % (  # pragma: no cover
            self.__class__.__name__,
            self.header.identifier,
            repr(self.objects).replace("\n", " ").replace("  ", " "),
        )

    @classmethod
    def from_buffer(cls, buf, filename=None):
        archive_info, payload = get_archive_info_and_remainder(buf)
        if not repr(archive_info):
            raise ValueError(
                "Segment doesn't seem to start with an ArchiveInfo!"
            )  # pragma: no cover

        payloads = []

        n = 0
        for message_info in archive_info.message_infos:
            try:
                if message_info.type == 0 and archive_info.should_merge and payloads:
                    base_message = archive_info.message_infos[
                        message_info.base_message_index
                    ]
                    klass = partial(
                        ProtobufPatch.FromString,
                        message_info,
                        ID_NAME_MAP[base_message.type],
                    )
                else:
                    klass = ID_NAME_MAP[message_info.type]
            except KeyError:  # pragma: no cover
                raise NotImplementedError(
                    "Don't know how to parse Protobuf message type "
                    + str(message_info.type)
                )
            try:
                message_payload = payload[n : n + message_info.length]
                if hasattr(klass, "FromString"):
                    output = klass.FromString(message_payload)
                else:
                    output = klass(message_payload)
            except Exception as e:  # pragma: no cover
                raise ValueError(
                    "Failed to deserialize %s payload of length %d: %s"
                    % (klass, message_info.length, e)
                )
            payloads.append(output)
            n += message_info.length

        return cls(archive_info, payloads), payload[n:]

    def to_dict(self):
        return {
            "header": header_to_dict(self.header),
            "objects": [message_to_dict(message) for message in self.objects],
        }


def message_to_dict(message):
    if hasattr(message, "to_dict"):
        return message.to_dict()
    output = MessageToDict(message)
    output["_pbtype"] = type(message).DESCRIPTOR.full_name
    return output


def header_to_dict(message):
    output = message_to_dict(message)
    for message_info in output["messageInfos"]:
        del message_info["length"]
    return output


def get_archive_info_and_remainder(buf):
    msg_len, new_pos = _DecodeVarint32(buf, 0)
    n = new_pos
    msg_buf = buf[n : n + msg_len]
    n += msg_len
    return ArchiveInfo.FromString(msg_buf), buf[n:]
