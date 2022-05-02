from numbers_parser.file import read_numbers_file


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
