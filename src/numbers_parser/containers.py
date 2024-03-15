import math
import re
from pathlib import Path

from numbers_parser.constants import PACKAGE_ID, SUPPORTED_NUMBERS_VERSIONS
from numbers_parser.iwafile import IWAFile, copy_object_to_iwa_file, create_iwa_segment
from numbers_parser.iwork import IWork, IWorkHandler


class ItemsList:
    def __init__(self, model, refs, item_class):
        self._item_name = item_class.__name__.lower()
        self._items = [item_class(model, id) for id in refs]

    def __getitem__(self, key: int):
        if isinstance(key, int):
            if key < 0:
                key += len(self._items)
            if key >= len(self._items):
                raise IndexError(f"index {key} out of range")
            return self._items[key]
        elif isinstance(key, str):
            for item in self._items:
                if item.name == key:
                    return item
            raise KeyError(f"no {self._item_name} named '{key}'")
        else:
            t = type(key).__name__
            raise LookupError(f"invalid index type {t}")

    def __len__(self) -> int:
        return len(self._items)

    def __contains__(self, key):
        return key.lower() in [x.name.lower() for x in self._items]

    def append(self, item):
        self._items.append(item)


class ObjectStore(IWorkHandler):
    def __init__(self, filepath: Path) -> int:
        self._objects = {}
        self._file_store = {}
        self._object_to_filename_map = {}
        self._dirty = {}
        self._iwork = IWork(handler=self)
        self._iwork.open(filepath)
        # TODO: why not just use the next available ID, i.e. without the offset?
        self._max_id = max(self._objects.keys())
        self._max_id = math.ceil(self._max_id / 1000000) * 1000000

    def save(self, filepath: Path, package: bool) -> None:
        self._iwork.save(filepath, self._file_store, package)

    def store_object(self, filename: str, identifier: int, archive: object) -> None:
        self._objects[identifier] = archive
        self._object_to_filename_map[identifier] = filename

    def store_file(self, filename: str, blob: bytes) -> None:
        self._file_store[filename] = blob

    def allowed_format(self, extension: str) -> bool:
        """bool: Return ``True`` if the filename extension is supported by the handler."""
        return True if extension == ".numbers" else False

    def allowed_version(self, version: str) -> bool:
        """bool: Return ``True`` if the document version is allowed."""
        version = re.sub(r"(\d+)\.(\d+)\.\d+", r"\1.\2", version)
        return True if version in SUPPORTED_NUMBERS_VERSIONS else False

    def new_message_id(self):
        """Return the next available message ID for object creation."""
        self._max_id += 1
        self._objects[PACKAGE_ID].last_object_identifier = self._max_id
        return self._max_id

    def create_object_from_dict(self, iwa_file: str, object_dict: dict, cls: object, append=False):
        """Create a new object and store the associated IWA segment. Return the
        message ID for the object and the newly created object. If the IWA
        file cannot be found, it will be created.
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
            self._file_store[iwa_pathname].chunks[0].archives.append(iwa_segment)

        self._objects[new_id] = cls(**object_dict)
        self._object_to_filename_map[new_id] = iwa_pathname
        return new_id, self._objects[new_id]

    def update_object_file_store(self):
        """Copy the protobuf messages from any updated object to the cached
        version in the file store so this can be saved to a new document.
        """
        for obj_id in self._objects:
            copy_object_to_iwa_file(
                self._file_store[self._object_to_filename_map[obj_id]],
                self._objects[obj_id],
                obj_id,
            )

    @property
    def file_store(self):
        return self._file_store

    def __getitem__(self, key: str):
        return self._objects[key]

    def __contains__(self, key: str):
        return key in self._objects

    def __len__(self) -> int:
        return len(self._objects)

    # Don't cache: new tables and sheets can be added at runtime
    def find_refs(self, ref_name) -> list:
        return [k for k, v in self._objects.items() if type(v).__name__ == ref_name]
