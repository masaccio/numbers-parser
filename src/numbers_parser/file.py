import logging
import os
from io import BytesIO
from sys import version_info
from zipfile import BadZipFile, ZipFile

from numbers_parser.exceptions import FileError, FileFormatError
from numbers_parser.iwafile import IWAFile, is_iwa_file

logger = logging.getLogger(__name__)
debug = logger.debug


def open_zipfile(file):
    """Open Zip file with the correct filename encoding supported by current python"""
    # Coverage is python version dependent, so one path with always fail coverage
    if version_info.minor >= 11:  # pragma: no cover
        return ZipFile(file, metadata_encoding="utf-8")
    else:  # pragma: no cover
        return ZipFile(file)


def read_numbers_file(path, file_handler, object_handler=None):
    debug("read_numbers_file: path=%s", path)
    if os.path.isdir(path):
        if os.path.isfile(os.path.join(path, "Index.zip")):
            get_objects_from_zip_file(os.path.join(path, "Index.zip"), file_handler, object_handler)
        else:
            for filename in os.listdir(path):
                filepath = os.path.join(path, filename)
                if os.path.isdir(filepath):
                    read_numbers_file(filepath, file_handler, object_handler)
                else:
                    f = open(filepath, "rb")
                    if filename.endswith(".iwa"):
                        blob = f.read()
                        extract_iwa_archives(blob, filepath, file_handler, object_handler)
                    blob = f.read()
                    file_handler(os.path.join(path, filename), blob)
    else:
        try:
            zipf = open_zipfile(path)
        except BadZipFile:
            raise FileFormatError("Invalid Numbers file") from None
        except FileNotFoundError:
            raise FileError("No such file or directory") from None

        try:
            index_zip = [f for f in zipf.namelist() if f.lower().endswith("index.zip")]
            if len(index_zip) > 0:
                index_data = BytesIO(zipf.read(index_zip[0]))
                get_objects_from_zip_stream(open_zipfile(index_data), file_handler, object_handler)
            else:
                get_objects_from_zip_stream(zipf, file_handler, object_handler)
        except BadZipFile:
            raise FileFormatError("Invalid Numbers file") from None


def write_numbers_file(filename, file_store):
    zipf = ZipFile(filename, "w")
    for filename, blob in file_store.items():
        if isinstance(blob, IWAFile):
            zipf.writestr(filename, blob.to_buffer())
        else:
            zipf.writestr(filename, blob)
    zipf.close()


def get_objects_from_zip_file(path, file_handler, object_handler):
    try:
        zipf = open_zipfile(path)
    except BadZipFile:
        raise FileFormatError("Invalid Numbers file") from None

    get_objects_from_zip_stream(zipf, file_handler, object_handler)


def get_objects_from_zip_stream(zipf, file_handler, object_handler):
    for filename in zipf.namelist():
        if filename.endswith(".iwa"):
            blob = zipf.read(filename)
            extract_iwa_archives(blob, filename, file_handler, object_handler)
        else:
            blob = zipf.read(filename)
            file_handler(filename, blob)


def extract_iwa_archives(blob, filename, file_handler, object_handler):
    if not is_iwa_file(blob):
        return

    try:
        debug("extract_iwa_archives: filename=%s", filename)
        iwaf = IWAFile.from_buffer(blob, filename)
    except Exception as e:
        raise FileFormatError(f"{filename}: invalid IWA file {filename}") from e

    if object_handler is not None:
        # Data from Numbers always has just one chunk. Some archives
        # have multiple objects though they appear not to contain
        # useful data.
        for archive in iwaf.chunks[0].archives:
            identifier = archive.header.identifier
            object_handler(identifier, archive.objects[0], filename)

    file_handler(filename, iwaf)
