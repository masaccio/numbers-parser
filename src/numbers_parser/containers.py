from zipfile import ZipFile, BadZipFile
from numbers_parser.codec import IWAFile
from pathlib import Path
from io import BytesIO

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
    def __init__(self, path: str) -> int:
        self._object_store = {}
        if os.path.isdir(path):
            if os.path.isfile(os.path.join(path, "Index.zip")):
                self._get_objects_from_zip_file(os.path.join(path, "Index.zip"))
            else:
                iwa_files = list(Path(path).rglob("*.iwa"))
                for iwa_filename in iwa_files:
                    f = open(iwa_filename, "rb")
                    contents = f.read()
                    self._extract_iwa_archives(contents, iwa_filename)
        else:
            try:
                zipf = ZipFile(path)
            except BadZipFile as e: # pragma: no cover
                raise FileError(f"{path}: " + str(e))

            if "Index.zip" in zipf.namelist():
                index_data = BytesIO(zipf.read("Index.zip"))
                self._get_objects_from_zip_stream(ZipFile(index_data))
            else:
                self._get_objects_from_zip_stream(zipf)

    def __getitem__(self, key: str):
        return self._object_store[key]

    def __len__(self) -> int:
        return len(self._object_store)

    def _get_objects_from_zip_file(self, path):
        try:
            zipf = ZipFile(path)
        except BadZipFile as e: # pragma: no cover
            raise FileError(f"{path}: " + str(e))

        self._get_objects_from_zip_stream(zipf)

    def _get_objects_from_zip_stream(self, zipf):
        iwa_files = filter(lambda x: x.endswith(".iwa"), zipf.namelist())
        for iwa_filename in iwa_files:
            # TODO: LZFSE compressed according to /usr/bin/file
            if "OperationStorage" in iwa_filename:
                continue
            contents = zipf.read(iwa_filename)
            self._extract_iwa_archives(contents, iwa_filename)


    def find_refs(self, ref_name) -> list:
        refs = [k for k, v in self._object_store.items() if type(v).__name__ == ref_name]
        return refs


    def _extract_iwa_archives(self, contents, iwa_filename):
        objects = {}
        try:
            iwaf = IWAFile.from_buffer(contents, iwa_filename)
        except Exception as e:  # pragma: no cover
            raise FileFormatError(f"{iwa_filename}: invalid IWA file {iwa_filename}") from e

        if len(iwaf.chunks) != 1:
            raise FileFormatError(f"{iwa_filename}: chunk count != 1 in {iwa_filename}") # pragma: no cover
        for archive in iwaf.chunks[0].archives:
            if len(archive.objects) == 0:
                raise FileFormatError(f"{iwa_filename}: no objects in {iwa_filename}") # pragma: no cover

            identifier = archive.header.identifier
            if identifier in objects:
                raise FileFormatError(f"{iwa_filename}: duplicate reference {identifier}") # pragma: no cover

            # TODO: what should we do for len(archive.objects) > 1?
            self._object_store[identifier] = archive.objects[0]
