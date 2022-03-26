from zipfile import ZipFile, BadZipFile
from numbers_parser.iwafile import IWAFile
from numbers_parser.exceptions import FileError, FileFormatError
from io import BytesIO

import os


def read_numbers_file(path, handler, store_objects=True):
    if os.path.isdir(path):
        if os.path.isfile(os.path.join(path, "Index.zip")):
            _get_objects_from_zip_file(
                os.path.join(path, "Index.zip"), handler, store_objects
            )
        else:
            for filename in os.listdir(path):
                filepath = os.path.join(path, filename)
                if os.path.isdir(filepath):
                    read_numbers_file(filepath, handler, store_objects)
                else:
                    f = open(filepath, "rb")
                    if filename.endswith(".iwa"):
                        contents = f.read()
                        _extract_iwa_archives(
                            contents, filepath, handler, store_objects
                        )
                    elif not store_objects:
                        contents = f.read()
                        handler(contents, os.path.join(path, filename))
    else:
        try:
            zipf = ZipFile(path)
        except BadZipFile as e:  # pragma: no cover
            raise FileError(f"{path}: " + str(e))

        if "Index.zip" in zipf.namelist():
            index_data = BytesIO(zipf.read("Index.zip"))
            _get_objects_from_zip_stream(ZipFile(index_data), handler, store_objects)
        else:
            _get_objects_from_zip_stream(zipf, handler, store_objects)


def _get_objects_from_zip_file(path, handler, store_objects=True):
    try:
        zipf = ZipFile(path)
    except BadZipFile as e:  # pragma: no cover
        raise FileError(f"{path}: " + str(e))

    _get_objects_from_zip_stream(zipf, handler, store_objects)


def _get_objects_from_zip_stream(zipf, handler, store_objects):
    for filename in zipf.namelist():
        if filename.endswith(".iwa"):
            contents = zipf.read(filename)
            _extract_iwa_archives(contents, filename, handler, store_objects)
        elif not store_objects:
            handler(zipf.read(filename), filename)


def _extract_iwa_archives(contents, filename, handler, store_objects):
    # TODO: LZFSE compressed according to /usr/bin/file
    if "OperationStorage" in filename:
        return

    try:
        iwaf = IWAFile.from_buffer(contents, filename)
    except Exception as e:  # pragma: no cover
        raise FileFormatError(f"{filename}: invalid IWA file {filename}") from e

    if store_objects:
        if len(iwaf.chunks) != 1:
            raise FileFormatError(
                f"{filename}: chunk count != 1 in {filename}"
            )  # pragma: no cover
        for archive in iwaf.chunks[0].archives:
            if len(archive.objects) == 0:
                raise FileFormatError(
                    f"{filename}: no objects in {filename}"
                )  # pragma: no cover

            identifier = archive.header.identifier
            # TODO: what should we do for len(archive.objects) > 1?
            handler(identifier, archive.objects[0])
    else:
        handler(iwaf, filename)
