from zipfile import ZipFile, BadZipFile
from numbers_parser.codec import IWAFile

class NumbersError(Exception):
    """Base class for other exceptions"""

    pass


class FileError(NumbersError):
    """Raised for IO and other OS errors"""

    pass


class FileFormatError(NumbersError):
    """Raised for parsing errors during file load"""

    pass


class ItemsList:
    def __init__(self, objects, refs, item_class):
        self._item_name = item_class.__name__.lower()
        self._items = [item_class(objects, _) for _ in refs]

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

    def __len__(self):
        return len(self._items)


class ObjectStore:
    def __init__(self, filename):
        try:
            zipf = ZipFile(filename)
        except BadZipFile as e:
            raise FileError(f"{filename}: " + str(e))

        objects = {}
        iwa_files = filter(lambda x: x.endswith(".iwa"), zipf.namelist())
        for iwa_filename in iwa_files:
            contents = zipf.read(iwa_filename)

            try:
                iwaf = IWAFile.from_buffer(contents, iwa_filename)
            except Error as e:
                raise FileFormatError(f"{filename}: invalid IWA file {iwa_filename}") from e

            if len(iwaf.chunks) != 1:
                raise FileFormatError(f"{filename}: chunk count != 1 in {iwa_filename}")
            for archive in iwaf.chunks[0].archives:
                if len(archive.objects) == 0:
                    raise FileFormatError(f"{filename}: no objects in {iwa_filename}")

                identifier = archive.header.identifier
                if identifier in objects:
                    raise FileFormatError(f"{filename}: duplicate reference {identifier}")

                if len(archive.objects) == 1:
                    objects[identifier] = archive.objects[0]
                else:
                    # print(f"warning: {iwa_filename}: found", len(archive.objects), "objects")
                    pass

        self._object_store = objects

    def __getitem__(self, key):
        return self._object_store[key]

    def __len__(self):
        return len(self._object_store)

    def find_refs(self, ref_name):
        refs = [k for k, v in self._object_store.items() if type(v).__name__ == ref_name]
        return refs

    def find_objects(self, ref_name, class_name):
        refs = self.find_refs(ref_name)
        class_ = getattr(importlib.import_module(__name__), class_name)
        return [class_(self, obj_id) for obj_id in refs]
