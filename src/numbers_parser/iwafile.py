# Forked from https://github.com/psobot/keynote-parser/blob/master/keynote_parser/codec.py

import logging
import struct
from functools import partial
from struct import unpack

import snappy
from google.protobuf.internal.decoder import _DecodeVarint32
from google.protobuf.internal.encoder import _VarintBytes
from google.protobuf.json_format import MessageToDict, ParseDict
from google.protobuf.message import EncodeError

from numbers_parser.exceptions import NotImplementedError
from numbers_parser.generated.mapping import ID_NAME_MAP, NAME_CLASS_MAP, NAME_ID_MAP
from numbers_parser.generated.TSPArchiveMessages_pb2 import ArchiveInfo

logger = logging.getLogger(__name__)
debug = logger.debug


class IWAFile:
    def __init__(self, chunks, filename=None) -> None:
        self.chunks = chunks
        self.filename = filename

    @classmethod
    def from_buffer(cls, data, filename=None):
        try:
            chunks = []
            while data:
                debug("from_buffer: filename=%s len=%d", filename, len(data))
                chunk, data = IWACompressedChunk.from_buffer(data, filename)
                chunks.append(chunk)

            return cls(chunks, filename)
        except Exception as e:  # pragma: no cover
            if filename:
                raise ValueError("Failed to deserialize " + filename) from e
            raise

    @classmethod
    def from_dict(cls, _dict):
        return cls([IWACompressedChunk.from_dict(chunk) for chunk in _dict["chunks"]])

    def to_dict(self):
        try:
            return {"chunks": [chunk.to_dict() for chunk in self.chunks]}
        except Exception as e:  # pragma: no cover
            if self.filename:
                raise ValueError("Failed to serialize " + self.filename) from e
            raise

    def to_buffer(self):
        return b"".join([chunk.to_buffer() for chunk in self.chunks])


class IWACompressedChunk:
    def __init__(self, archives) -> None:
        self.archives = archives

    def __eq__(self, other):
        return self.archives == other.archives  # pragma: no cover

    @classmethod
    def _decompress_all(cls, data):
        while data:
            header = data[:4]

            first_byte = header[0]
            if first_byte != 0x00:  # pragma: no cover
                msg = f"IWA chunk does not start with 0x00! (found {first_byte:x})"
                raise ValueError(
                    msg,
                )

            length = unpack("<I", bytes(header[1:]) + b"\x00")[0]
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
            debug("from_buffer: filename=%s len=%d", filename, len(data))
            archive, data = IWAArchiveSegment.from_buffer(data, filename)
            archives.append(archive)
        return cls(archives), None

    @classmethod
    def from_dict(cls, _dict):
        return cls([IWAArchiveSegment.from_dict(d) for d in _dict["archives"]])

    def to_dict(self):
        return {"archives": [archive.to_dict() for archive in self.archives]}

    def to_buffer(self):
        uncompressed = b"".join([archive.to_buffer() for archive in self.archives])
        payloads = []
        while uncompressed:
            payloads.append(snappy.compress(uncompressed[:65536]))
            uncompressed = uncompressed[65536:]
        return b"".join(
            [b"\x00" + struct.pack("<I", len(payload))[:3] + payload for payload in payloads],
        )


class ProtobufPatch:
    def __init__(self, data) -> None:
        self.data = data

    def __eq__(self, other):
        return self.data == other.data  # pragma: no cover

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__} {self.data}>"  # pragma: no cover

    def to_dict(self):
        return message_to_dict(self.data)

    @classmethod
    def FromString(cls, message_info, proto_klass, data):  # noqa: N802
        # Note versus Peter Sobot's implementation: we can ignore some of
        # the unimplemented patching of Protobufs. Specifically deserializing
        # when len(diff_field_path) > 1 or when fields_to_remove is present.
        return cls(proto_klass.FromString(data))

    def SerializeToString(self):  # noqa: N802
        return self.data.SerializePartialToString()


class IWAArchiveSegment:
    def __init__(self, header, objects) -> None:
        self.header = header
        self.objects = objects

    def __eq__(self, other):
        return self.header == other.header and self.objects == other.objects  # pragma: no cover

    def __repr__(self) -> str:  # pragma: no cover
        self_str = repr(self.objects).replace("\n", " ").replace("  ", " ")
        return f"<{self.__class__.__name__} identifier={self.header.identifier} objects={self_str}>"

    @classmethod
    def from_buffer(cls, buf, filename=None):
        archive_info, payload = get_archive_info_and_remainder(buf)
        if not repr(archive_info):
            msg = "Segment doesn't seem to start with an ArchiveInfo!"
            raise ValueError(
                msg,
            )  # pragma: no cover

        payloads = []

        n = 0
        for message_info in archive_info.message_infos:
            try:
                if message_info.type == 0 and archive_info.should_merge and payloads:
                    base_message = archive_info.message_infos[message_info.base_message_index]
                    klass = partial(
                        ProtobufPatch.FromString,
                        message_info,
                        ID_NAME_MAP[base_message.type],
                    )
                else:
                    klass = ID_NAME_MAP[message_info.type]
            except KeyError:  # pragma: no cover
                raise NotImplementedError(
                    "Don't know how to parse Protobuf message type " + str(message_info.type),
                ) from None
            try:
                message_payload = payload[n : n + message_info.length]
                if hasattr(klass, "FromString"):
                    output = klass.FromString(message_payload)
                else:
                    output = klass(message_payload)
            except Exception as e:  # pragma: no cover
                raise ValueError(
                    "Failed to deserialize %s payload of length %d: %s"
                    % (klass, message_info.length, e),
                ) from None
            payloads.append(output)
            n += message_info.length

        return cls(archive_info, payloads), payload[n:]

    @classmethod
    def from_dict(cls, _dict):
        header = dict_to_header(_dict["header"])
        objects = []
        for _message_info, o in zip(header.message_infos, _dict["objects"]):
            objects.append(dict_to_message(o))
        return cls(header, objects)

    def to_dict(self):
        return {
            "header": header_to_dict(self.header),
            "objects": [message_to_dict(message) for message in self.objects],
        }

    def to_buffer(self):
        # Each message_info as part of the header needs to be updated
        # so that its length matches the object contained within.
        for obj, message_info in zip(self.objects, self.header.message_infos):
            try:
                object_length = len(obj.SerializeToString())
                provided_length = message_info.length
                if object_length != provided_length:
                    message_info.length = object_length
            except EncodeError as e:  # pragma: no cover  # noqa: PERF203
                msg = (
                    f"Failed to encode object: {e}\nObject: '{obj!r}'\nMessage info: {message_info}"
                )
                raise ValueError(
                    msg,
                ) from None
        return b"".join(
            [_VarintBytes(self.header.ByteSize()), self.header.SerializeToString()]
            + [obj.SerializeToString() for obj in self.objects],
        )


def message_to_dict(message):
    if hasattr(message, "to_dict"):
        return message.to_dict()
    output = MessageToDict(message, preserving_proto_field_name=True)
    output["_pbtype"] = type(message).DESCRIPTOR.full_name
    return output


def header_to_dict(message):
    output = message_to_dict(message)
    for message_info in output["message_infos"]:
        del message_info["length"]
    return output


def dict_to_message(_dict):
    _type = _dict["_pbtype"]
    del _dict["_pbtype"]
    return ParseDict(_dict, NAME_CLASS_MAP[_type](), ignore_unknown_fields=True)


def dict_to_header(_dict):
    for message_info in _dict["message_infos"]:
        # set a dummy length value that we'll overwrite later
        message_info["length"] = 0
    return dict_to_message(_dict)


def get_archive_info_and_remainder(buf):
    msg_len, new_pos = _DecodeVarint32(buf, 0)
    n = new_pos
    msg_buf = buf[n : n + msg_len]
    n += msg_len
    return ArchiveInfo.FromString(msg_buf), buf[n:]


def create_iwa_segment(obj_id: int, cls: object, object_dict: dict) -> object:
    full_name = cls.DESCRIPTOR.full_name
    type_id = NAME_ID_MAP[full_name]
    header = {
        "_pbtype": "TSP.ArchiveInfo",
        "identifier": str(obj_id),
        "message_infos": [
            {
                "type": type_id,
                "version": [1, 0, 5],
            },
        ],
    }
    object_dict["_pbtype"] = full_name
    archive_dict = {"header": header, "objects": [object_dict]}

    return IWAArchiveSegment.from_dict(archive_dict)


def find_references(obj, references=list) -> None:
    if not hasattr(obj, "DESCRIPTOR"):
        return
    if type(obj).__name__ == "Reference":
        references.append(obj.identifier)
        return
    for field_desc in obj.ListFields():
        _, field = field_desc
        if type(field).__name__ == "Reference":
            references.append(field.identifier)
        elif "Repeated" in type(field).__name__:
            for item in field:
                find_references(item, references)
        elif hasattr(field, "DESCRIPTOR"):
            find_references(field, references)


def copy_object_to_iwa_file(iwa_file: IWAFile, obj: object, obj_id: int) -> None:
    for archive in iwa_file.chunks[0].archives:
        if archive.header.identifier == obj_id:
            archive.objects[0].CopyFrom(obj)
            references = []
            find_references(archive.objects[0], references)
            if len(references) > 0:
                msg_info = archive.header.message_infos[0]
                while len(msg_info.object_references) > 0:
                    _ = msg_info.object_references.pop()
                for reference in references:
                    msg_info.object_references.append(reference)


def is_iwa_file(data):
    data_length = len(data)
    length = 0
    while data:
        header = data[:4]

        first_byte = header[0]
        if first_byte != 0x00:
            return False

        segment_length = unpack("<I", bytes(header[1:]) + b"\x00")[0]
        length += segment_length + 4
        data = data[4 + segment_length :]
    return length == data_length


def extensions(obj) -> list[object]:
    return [obj.Extensions[field] for field, _ in obj.ListFields() if field.is_extension]


def find_extension(obj, name: str) -> object:
    all_extensions = extensions(obj)
    filtered = [getattr(x, name) for x in all_extensions if getattr(x, name, None) is not None]
    return filtered[0]
