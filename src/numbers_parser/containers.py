from numbers_parser.file import read_numbers_file
from numbers_parser.iwafile import create_iwa_segment


class ItemsList:
    def __init__(self, model, refs, item_class):
        self._item_name = item_class.__name__.lower()
        self._items = [item_class(model, _) for _ in refs]

    def __getitem__(self, key: int):
        if type(key) == int:
            if key < 0 or key >= len(self._items):
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


class ObjectStore:
    def __init__(self, path: str) -> int:
        self._objects = {}
        self._file_store = {}
        read_numbers_file(
            path,
            file_handler=lambda filename, blob: self._store_file(filename, blob),
            object_handler=lambda identifier, obj: self._store_object(identifier, obj),
        )

    def _new_message_id(self):
        """Return the next available message ID for object creation"""
        max_id = max(self._objects.keys())
        return max_id + 1

    def create_object_from_dict(self, iwa_file: str, object_dict: dict, cls: object):
        """Create a new object and store the associated IWA segment. Return the
        message ID for the object and the newly created object"""
        paths = [k for k, v in self._file_store.items() if iwa_file in k]
        if len(paths) == 0:
            raise LookupError(f"no IWA filename matching {iwa_file}")
        iwa_pathname = paths[0]

        new_id = self._new_message_id()
        iwa_segment = create_iwa_segment(new_id, cls, object_dict)

        self._file_store[iwa_pathname].chunks[0].archives.append(iwa_segment)
        self._objects[new_id] = cls(**object_dict)
        return new_id, self._objects[new_id]

    def update_object(self, obj_id, obj: object):
        """Copy the protobuf messages from the passed object to the cached
        version in the file store so this can be saved to a new document"""
        for filename, blob in self._file_store.items():
            for archive in blob.chunks[0].archives:
                if archive.header.identifier == obj_id:
                    archive.objects[0].CopyFrom(obj)
                    return

        raise LookupError(f"Cannot loate object for update with ID {obj_id}")

    def _store_object(self, identifier, obj):
        self._objects[identifier] = obj

    def _store_file(self, filename, blob):
        self._file_store[filename] = blob

    @property
    def file_store(self):
        return self._file_store

    def __getitem__(self, key: str):
        return self._objects[key]

    def __len__(self) -> int:
        return len(self._objects)

    def find_refs(self, ref_name) -> list:
        refs = [k for k, v in self._objects.items() if type(v).__name__ == ref_name]
        return refs
