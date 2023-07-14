import math

from numbers_parser.file import read_numbers_file
from numbers_parser.constants import PACKAGE_ID
from numbers_parser.iwafile import create_iwa_segment, copy_object_to_iwa_file, IWAFile


class ItemsList:
    def __init__(self, model, refs, item_class):
        self._item_name = item_class.__name__.lower()
        self._items = [item_class(model, _) for _ in refs]

    def __getitem__(self, key: int):
        if type(key) == int:
            if key < 0:
                key += len(self._items)
            if key >= len(self._items):
                raise IndexError(f"index {key} out of range")
            return self._items[key]
        elif type(key) == str:
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


class ObjectStore:
    def __init__(self, path: str) -> int:
        self._objects = {}
        self._file_store = {}
        self._object_to_filename_map = {}
        self._dirty = {}
        read_numbers_file(
            path,
            file_handler=lambda filename, blob: self._store_file(filename, blob),
            object_handler=lambda identifier, obj, filename: self._store_object(
                identifier, obj, filename
            ),
        )
        # TODO: why not just use the next available ID, i.e. without the offset?
        self._max_id = max(self._objects.keys())
        self._max_id = math.ceil(self._max_id / 1000000) * 1000000

    def new_message_id(self):
        """Return the next available message ID for object creation"""
        self._max_id += 1
        self._objects[PACKAGE_ID].last_object_identifier = self._max_id
        return self._max_id

    def create_object_from_dict(self, iwa_file: str, object_dict: dict, cls: object):
        """Create a new object and store the associated IWA segment. Return the
        message ID for the object and the newly created object. If the IWA
        file cannot be found, it will be created."""
        paths = [k for k, v in self._file_store.items() if iwa_file in k]
        if len(paths) == 0:
            iwa_pathname = None
        else:
            iwa_pathname = paths[0]

        new_id = self.new_message_id()
        iwa_segment = create_iwa_segment(new_id, cls, object_dict)

        if iwa_pathname is None:
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
        version in the file store so this can be saved to a new document"""
        for obj_id in self._objects.keys():
            copy_object_to_iwa_file(
                self._file_store[self._object_to_filename_map[obj_id]],
                self._objects[obj_id],
                obj_id,
            )

    def _store_object(self, identifier, obj, filename):
        self._objects[identifier] = obj
        self._object_to_filename_map[identifier] = filename

    def _store_file(self, filename, blob):
        self._file_store[filename] = blob

    @property
    def file_store(self):
        return self._file_store

    def __getitem__(self, key: str):
        return self._objects[key]

    def __len__(self) -> int:
        return len(self._objects)

    # Don't cache: new tables and sheets can be added at runtime
    def find_refs(self, ref_name) -> list:
        refs = [k for k, v in self._objects.items() if type(v).__name__ == ref_name]
        return refs
