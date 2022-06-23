from zipfile import ZipFile, BadZipFile
from numbers_parser.iwafile import IWAFile
from numbers_parser.exceptions import FileError, FileFormatError
from io import BytesIO

import os


def read_numbers_file(path, file_handler=None, object_handler=None):
    if os.path.isdir(path):
        if os.path.isfile(os.path.join(path, "Index.zip")):
            _get_objects_from_zip_file(
                os.path.join(path, "Index.zip"), file_handler, object_handler
            )
        else:
            for filename in os.listdir(path):
                filepath = os.path.join(path, filename)
                if os.path.isdir(filepath):
                    read_numbers_file(filepath, file_handler, object_handler)
                else:
                    f = open(filepath, "rb")
                    if filename.endswith(".iwa"):
                        blob = f.read()
                        _extract_iwa_archives(
                            blob, filepath, file_handler, object_handler
                        )
                    if file_handler is not None:
                        blob = f.read()
                        file_handler(os.path.join(path, filename), blob)
    else:
        try:
            zipf = ZipFile(path)
        except BadZipFile as e:  # pragma: no cover
            raise FileError(f"{path}: " + str(e))

        index_zip = [f for f in zipf.namelist() if f.lower().endswith("index.zip")]
        if len(index_zip) > 0:
            index_data = BytesIO(zipf.read(index_zip[0]))
            _get_objects_from_zip_stream(
                ZipFile(index_data), file_handler, object_handler
            )
        else:
            _get_objects_from_zip_stream(zipf, file_handler, object_handler)


def write_numbers_file(filename, file_store):
    zipf = ZipFile(filename, "w")
    for filename, blob in file_store.items():
        if isinstance(blob, IWAFile):
            zipf.writestr(filename, blob.to_buffer())
        else:
            zipf.writestr(filename, blob)
    zipf.close()


def _get_objects_from_zip_file(path, file_handler, object_handler):
    try:
        zipf = ZipFile(path)
    except BadZipFile as e:  # pragma: no cover
        raise FileError(f"{path}: " + str(e))

    _get_objects_from_zip_stream(zipf, file_handler, object_handler)


def _get_objects_from_zip_stream(zipf, file_handler, object_handler):
    for filename in zipf.namelist():
        if filename.endswith(".iwa"):
            blob = zipf.read(filename)
            _extract_iwa_archives(blob, filename, file_handler, object_handler)
        elif file_handler is not None:
            blob = zipf.read(filename)
            file_handler(filename, blob)


def _extract_iwa_archives(blob, filename, file_handler, object_handler):
    # LZFSE compressed according to /usr/bin/file
    if "OperationStorage" in filename:
        return

    try:
        iwaf = IWAFile.from_buffer(blob, filename)
    except Exception as e:  # pragma: no cover
        raise FileFormatError(f"{filename}: invalid IWA file {filename}") from e

    if object_handler is not None:
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
            object_handler(identifier, archive.objects[0], filename)

    if file_handler is not None:
        file_handler(filename, iwaf)
