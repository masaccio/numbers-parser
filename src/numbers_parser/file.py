import logging
import os
import plistlib
import re
import warnings
from io import BytesIO
from sys import version_info
from typing import TextIO
from zipfile import BadZipFile, ZipFile

from numbers_parser.constants import _SUPPORTED_NUMBERS_VERSIONS
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


def test_document_version(fp: TextIO, path: str) -> None:
    doc_properties = plistlib.load(fp)
    doc_version = doc_properties["fileFormatVersion"]
    if doc_version not in _SUPPORTED_NUMBERS_VERSIONS:
        warnings.warn(f"{path}: unsupported version {doc_version}", RuntimeWarning, stacklevel=2)


def read_numbers_file(path, file_handler, object_handler=None):
    if os.path.isdir(path):
        if not path.endswith(".numbers"):
            raise FileFormatError("invalid Numbers document (not a .numbers directory)")

        properties_plist = os.path.join(path, "Metadata/Properties.plist")
        try:
            fp = open(properties_plist, "rb")
        except OSError:
            raise FileFormatError("invalid Numbers document (missing files)") from None
        test_document_version(fp, path)
        fp.close()

    read_numbers_file_contents(path, file_handler, object_handler)


def read_numbers_file_contents(path, file_handler, object_handler=None):
    debug("read_numbers_file: path=%s", path)
    if os.path.isdir(path):
        for filename in os.listdir(path):
            filepath = os.path.join(path, filename)
            if os.path.isdir(filepath):
                read_numbers_file_contents(filepath, file_handler, object_handler)
            elif filename == "Index.zip":
                get_objects_from_zip_file(filepath, file_handler, object_handler)
            else:
                f = open(filepath, "rb")
                blob = f.read()
                if filename.endswith(".iwa"):
                    package_filepath = re.sub(r".*\.numbers/*", "", filepath)
                    extract_iwa_archives(blob, package_filepath, file_handler, object_handler)
                else:
                    package_filepath = os.path.join(re.sub(r".*\.numbers/*", "", path), filename)
                    file_handler(package_filepath, blob)
    else:
        try:
            zipf = open_zipfile(path)
        except BadZipFile:
            raise FileFormatError("invalid Numbers document") from None
        except FileNotFoundError:
            raise FileError("no such file or directory") from None

        try:
            index_zip = [f for f in zipf.namelist() if f.lower().endswith("index.zip")]
            if len(index_zip) > 0:
                index_data = BytesIO(zipf.read(index_zip[0]))
                get_objects_from_zip_stream(open_zipfile(index_data), file_handler, object_handler)
            else:
                get_objects_from_zip_stream(zipf, file_handler, object_handler)
        except BadZipFile:
            raise FileFormatError("invalid Numbers document") from None


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
        raise FileFormatError("invalid Numbers document") from None

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
