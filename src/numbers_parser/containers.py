from zipfile import ZipFile, BadZipFile
from numbers_parser.codec import IWAFile
from pathlib import Path

import os


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

    def __len__(self) -> int:
        return len(self._items)


class ObjectStore:
    def __init__(self, filename: str) -> int:
        objects = {}
        if os.path.isdir(filename):
            if os.path.isfile(os.path.join(filename, "Index.zip")):
                try:
                    zipf = ZipFile(os.path.join(filename, "Index.zip"))
                except BadZipFile as e:
                    raise FileError(f"{index_zip}: " + str(e))

                iwa_files = filter(lambda x: x.endswith(".iwa"), zipf.namelist())
                for iwa_filename in iwa_files:
                    # TODO: LZFSE compressed according to /usr/bin/file
                    if "OperationStorage" in iwa_filename:
                        continue
                    contents = zipf.read(iwa_filename)
                    objects.update(extract_iwa_archives(contents, iwa_filename))
            else:
                iwa_files = list(Path(filename).rglob("*.iwa"))
                for iwa_filename in iwa_files:
                    f = open(iwa_filename, "rb")
                    contents = f.read()
                    objects.update(extract_iwa_archives(contents, iwa_filename))
        else:
            try:
                zipf = ZipFile(filename)
            except BadZipFile as e:
                raise FileError(f"{filename}: " + str(e))

            iwa_files = filter(lambda x: x.endswith(".iwa"), zipf.namelist())
            for iwa_filename in iwa_files:
                contents = zipf.read(iwa_filename)
                objects.update(extract_iwa_archives(contents, iwa_filename))

        self._object_store = objects

    def __getitem__(self, key: str):
        return self._object_store[key]

    def __len__(self) -> int:
        return len(self._object_store)

    def find_refs(self, ref_name) -> list:
        refs = [
            k for k, v in self._object_store.items() if type(v).__name__ == ref_name
        ]
        return refs


def extract_iwa_archives(contents, iwa_filename):
    objects = {}
    try:
        iwaf = IWAFile.from_buffer(contents, iwa_filename)
    except Exception as e:
        raise FileFormatError(f"{iwa_filename}: invalid IWA file {iwa_filename}") from e

    if len(iwaf.chunks) != 1:
        raise FileFormatError(f"{iwa_filename}: chunk count != 1 in {iwa_filename}")
    for archive in iwaf.chunks[0].archives:
        if len(archive.objects) == 0:
            raise FileFormatError(f"{iwa_filename}: no objects in {iwa_filename}")

        identifier = archive.header.identifier
        if identifier in objects:
            raise FileFormatError(f"{iwa_filename}: duplicate reference {identifier}")

        if len(archive.objects) == 1:
            objects[identifier] = archive.objects[0]
        else:
            #  print(f"warning: {iwa_filename}: found", len(archive.objects), "objects")
            pass

    return objects
