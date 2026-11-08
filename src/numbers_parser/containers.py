"""Containers module helpers."""

import math
import re
from pathlib import Path
from typing import TYPE_CHECKING, Any, Generic, Protocol, TypeVar, cast, overload

from google.protobuf.message import Message as ProtoMessage

from numbers_parser.constants import DOCUMENT_ID, PACKAGE_ID, SUPPORTED_NUMBERS_VERSIONS
from numbers_parser.iwafile import IWAFile, copy_object_to_iwa_file, create_iwa_segment
from numbers_parser.iwork import IWork, IWorkHandler

if TYPE_CHECKING:
    import numbers_parser.generated.TSPArchiveMessages_pb2 as TSPArchiveMessage
    from numbers_parser.model import NumbersModel


_ProtoT = TypeVar("_ProtoT", bound=ProtoMessage)


class _NamedItem(Protocol):
    """Protocol defining the required name attribute for items stored in an ItemsList."""

    @property
    def name(self) -> str:
        """str: The name of the item."""


_ItemT = TypeVar("_ItemT", bound=_NamedItem)


class ItemsList(Generic[_ItemT]):
    """Provide indexed and named access to model-backed items."""

    def __init__(self, model: "NumbersModel", refs: list[int], item_class: type[_ItemT]) -> None:
        """Initialize the instance state."""
        self._item_name = item_class.__name__.lower()
        self._items = [item_class(model, x) for x in refs]  # type: ignore[call-arg]

    @overload
    def __getitem__(self, key: int) -> _ItemT: ...  # pragma: no cover

    @overload
    def __getitem__(self, key: str) -> _ItemT: ...  # pragma: no cover

    def __getitem__(self, key: int | str) -> _ItemT:
        """Return the item at the requested key or index."""
        if isinstance(key, int):
            if key < 0:
                key += len(self._items)
            if key < 0 or key >= len(self._items):
                msg = f"index {key} out of range"
                raise IndexError(msg)
            return self._items[key]
        if isinstance(key, str):
            for item in self._items:
                if item.name.lower() == key.lower():
                    return item
            msg = f"no {self._item_name} named '{key}'"
            raise KeyError(msg)
        t = type(key).__name__
        msg = f"invalid index type {t}"
        raise LookupError(msg)

    def __len__(self) -> int:
        """Return the number of contained items."""
        return len(self._items)

    def __contains__(self, key: str) -> bool:
        """Return whether the value is present."""
        return key.lower() in [x.name.lower() for x in self._items]

    def append(self, item: _ItemT) -> None:
        """Append the value."""
        self._items.append(item)


class ObjectStore(IWorkHandler):
    """Store and retrieve archive objects and their IWA files."""

    def __init__(self, filepath: Path, password: str | None) -> None:
        """Initialize the instance state."""
        self._objects: dict[int, ProtoMessage] = {}
        self._file_store: dict[str, IWAFile | bytes] = {}
        self._object_to_filename_map: dict[int, str] = {}
        self._dirty: dict[int, bool] = {}
        self._iwork = IWork(handler=self)
        self._iwork.open(filepath, password)
        # TODO: why not just use the next available ID, i.e. without the offset?
        self._max_id = max(self._objects.keys())
        self._max_id = math.ceil(self._max_id / 1000000) * 1000000

    def save(self, filepath: Path, package: bool, password: str | None, hint: str) -> None:
        """Save the value."""
        self._iwork.save(filepath, self._file_store, package, password, hint)

    def store_object(self, filename: str, identifier: int, archive: object) -> None:
        """Store object."""
        self._objects[identifier] = cast("ProtoMessage", archive)
        self._object_to_filename_map[identifier] = filename

    def store_file(self, filename: str, blob: bytes) -> None:
        """Store file."""
        self._file_store[filename] = blob

    def allowed_format(self, extension: str) -> bool:
        """bool: Return ``True`` if the filename extension is supported by the handler."""
        return extension == ".numbers"

    def allowed_version(self, version: str) -> bool:
        """bool: Return ``True`` if the document version is allowed."""
        version = re.sub(r"(\d+)\.(\d+)\.\d+", r"\1.\2", version)
        return version in SUPPORTED_NUMBERS_VERSIONS

    def new_message_id(self) -> int:
        """Return the next available message ID for object creation."""
        self._max_id += 1
        metadata = cast("TSPArchiveMessage.PackageMetadata", self._objects[PACKAGE_ID])
        metadata.last_object_identifier = self._max_id
        return self._max_id

    def create_object_from_dict(
        self,
        iwa_file: str,
        object_dict: dict[str, Any],
        cls: type[ProtoMessage],
        append: bool = False,
    ) -> tuple[int, ProtoMessage]:
        """
        Create and store an object with its IWA segment.

        Return the message ID and object. Create the IWA file if it is missing.
        """
        paths = [k for k, v in self._file_store.items() if iwa_file in k]
        iwa_pathname = None if len(paths) == 0 else paths[0]

        new_id = self.new_message_id()
        iwa_segment = create_iwa_segment(new_id, cls, object_dict)

        if iwa_pathname is None and not append:
            iwa_pathname = iwa_file.format(new_id) + ".iwa"
            chunks = {"chunks": [{"archives": [iwa_segment.to_dict()]}]}
            self._file_store[iwa_pathname] = IWAFile.from_dict(chunks)
        else:
            iwa_pathname = cast("str", iwa_pathname)
            iwa_obj = cast("IWAFile", self._file_store[iwa_pathname])
            iwa_obj.chunks[0].archives.append(iwa_segment)

        self._objects[new_id] = cls(**object_dict)
        self._object_to_filename_map[new_id] = iwa_pathname
        return new_id, self._objects[new_id]

    def update_object_file_store(self) -> None:
        """
        Update cached IWA files from modified protobuf messages.

        This keeps the file store ready for saving to a new document.
        """
        for obj_id, obj in self._objects.items():
            iwa_file = cast("IWAFile", self._file_store[self._object_to_filename_map[obj_id]])
            copy_object_to_iwa_file(iwa_file, obj, obj_id)

    def remove_unreferenced_objects(self) -> None:
        """Remove objects from the ObjectStore that are not referenced."""
        referenced_ids = {DOCUMENT_ID, PACKAGE_ID}

        def find_references(msg: ProtoMessage) -> None:
            """Recursively find all TSP.Reference messages in protobuf."""
            if hasattr(msg, "identifier") and msg.DESCRIPTOR.name == "Reference":
                referenced_ids.add(msg.identifier)

            for field_desc, value in msg.ListFields():
                if field_desc.type != field_desc.TYPE_MESSAGE:
                    continue
                if field_desc.is_repeated:
                    for item in value:
                        find_references(item)
                else:
                    find_references(value)

        for obj in self._objects.values():
            find_references(obj)

        metadata = cast("TSPArchiveMessage.PackageMetadata", self._objects[PACKAGE_ID])
        unreferenced_ids = set(self._objects.keys()) - referenced_ids
        unreferenced_ids -= {c.identifier for c in metadata.components}

        # Delete unreferenced archives. In principal we could delete unreferenced files,
        # but deleting tables/sheets is unsupported so this never happens.
        for obj_id in unreferenced_ids:
            del self._objects[obj_id]
            filename = self._object_to_filename_map.pop(obj_id)
            iwa_file = cast("IWAFile", self._file_store[filename])
            for chunk in iwa_file.chunks:
                chunk.archives = [
                    arch for arch in chunk.archives if arch.header.identifier != obj_id
                ]

    @property
    def file_store(self) -> dict[str, bytes | IWAFile]:
        """Return or update the IWA files associated with the object store."""
        return self._file_store

    def __getitem__(self, key: int) -> "ProtoMessage":
        """Return a generically typed object from the store."""
        return self._objects[key]

    def __contains__(self, key: int) -> bool:
        """Return whether the value is present."""
        return key in self._objects

    def __len__(self) -> int:
        """Return the number of contained items."""
        return len(self._objects)

    def get_as(
        self,
        key: int,
        expected_type: type[_ProtoT],  # noqa: ARG002, pylint: disable=unused-argument
    ) -> _ProtoT:
        """Fetch an object from the store and cast it to the expected Protobuf class."""
        return cast("_ProtoT", self._objects[key])

    def find_refs(self, ref_name: str) -> list[int]:
        """Find refs."""
        return [k for k, v in self._objects.items() if type(v).__name__ == ref_name]
