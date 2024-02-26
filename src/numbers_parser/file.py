import logging
import plistlib
import re
import warnings
from io import BytesIO
from pathlib import Path
from sys import version_info
from typing import Callable, List, Union
from zipfile import BadZipFile, ZipFile

from numbers_parser.constants import _SUPPORTED_NUMBERS_VERSIONS
from numbers_parser.exceptions import FileError, FileFormatError
from numbers_parser.iwafile import IWAFile, is_iwa_file

logger = logging.getLogger(__name__)
debug = logger.debug


def open_zipfile(filepath: Path):
    """Open Zip file with the correct filename encoding supported by current python"""
    # Coverage is python version dependent, so one path with always fail coverage
    if version_info.minor >= 11:  # pragma: no cover
        return ZipFile(filepath, metadata_encoding="utf-8")
    else:  # pragma: no cover
        return ZipFile(filepath)


def test_document_version(filepath: str, no_warn=False) -> None:
    """
    Test whether a document was created using a supported version of
    Numbers and raise an exception if not.
    """
    properties_plist = filepath / "Metadata/Properties.plist"
    try:
        with properties_plist.open(mode="rb") as fp:
            doc_properties = plistlib.load(fp)
            doc_version = doc_properties["fileFormatVersion"]
            if not no_warn and doc_version not in _SUPPORTED_NUMBERS_VERSIONS:
                warnings.warn(f"unsupported version {doc_version}", RuntimeWarning, stacklevel=2)
    except OSError:
        raise FileFormatError("invalid Numbers document (missing files)") from None


def read_numbers_file(filepath: Path, file_handler: Callable, object_handler: Callable = None):
    """
    Wrapper for the Numbers file reader that performs some tests on the validity of
    packages before reading them.
    """
    debug("read_numbers_file: path=%s", filepath)
    if filepath.is_dir():
        if filepath.suffix != ".numbers":
            raise FileFormatError("invalid Numbers document (not a .numbers directory)")
        test_document_version(filepath)

    read_numbers_file_contents(filepath, file_handler, object_handler)


def read_numbers_file_contents(
    filepath: Path, file_handler: Callable, object_handler: Callable = None
):
    """
    Read a Numbers file and iterate through all files and directories
    storing the files blobs and objects though the supplies callbacks.
    """
    if filepath.is_dir():
        for sub_filepath in filepath.iterdir():
            if sub_filepath.is_dir():
                read_numbers_file_contents(sub_filepath, file_handler, object_handler)
            elif sub_filepath.name.lower() == "index.zip":
                try:
                    zipf = open_zipfile(sub_filepath)
                except BadZipFile:
                    raise FileFormatError("invalid Numbers document") from None

                get_objects_from_zip_stream(zipf, file_handler, object_handler)
            else:
                with sub_filepath.open(mode="rb") as fh:
                    blob = fh.read()
                    package_filename = re.sub(r".*\.numbers/*", "", str(sub_filepath))
                    if sub_filepath.suffix == ".iwa":
                        extract_iwa_archives(blob, package_filename, file_handler, object_handler)
                    else:
                        file_handler(package_filename, blob)
    else:
        try:
            zipf = open_zipfile(filepath)
        except BadZipFile:
            raise FileFormatError("invalid Numbers document") from None
        except FileNotFoundError:
            raise FileError("no such file or directory") from None

        try:
            get_objects_from_zip_stream(zipf, file_handler, object_handler)
        except BadZipFile:
            raise FileFormatError("invalid Numbers document") from None


def write_numbers_file(filepath: Path, file_store: List[object], package: bool):
    if package:
        if filepath.is_dir():
            if not filepath.suffix == ".numbers":
                raise FileFormatError("invalid Numbers document (not a .numbers directory)")
            if not (filepath / "Index.zip").is_file():
                raise FileFormatError("folder is not a numbers package")
            test_document_version(filepath, no_warn=True)
        elif filepath.is_file():
            raise FileFormatError("cannot overwrite Numbers document file with package")
        else:
            filepath.mkdir()
        zipf = ZipFile(filepath / "Index.zip", "w")
        for blob_path, blob in file_store.items():
            if isinstance(blob, IWAFile):
                zipf.writestr(blob_path, blob.to_buffer())
            else:
                sub_filepath = filepath / blob_path
                if not sub_filepath.parent.is_dir():
                    sub_filepath.parent.mkdir()
                with sub_filepath.open(mode="wb") as fh:
                    fh.write(blob)
        zipf.close()
    else:
        zipf = ZipFile(filepath, "w")

        for filepath, blob in file_store.items():
            if isinstance(blob, IWAFile):
                zipf.writestr(filepath, blob.to_buffer())
            else:
                zipf.writestr(filepath, blob)
        zipf.close()


def get_objects_from_zip_stream(zipf, file_handler, object_handler):
    for filename in zipf.namelist():
        blob = zipf.read(filename)
        if filename.lower().endswith("index.zip"):
            index_data = BytesIO(blob)
            get_objects_from_zip_stream(open_zipfile(index_data), file_handler, object_handler)
        elif filename.endswith(".iwa"):
            extract_iwa_archives(blob, filename, file_handler, object_handler)
        else:
            file_handler(filename, blob)


def extract_iwa_archives(
    blob: bytes, filename: str, file_handler: Callable, object_handler: Union[Callable, None]
):
    """If blob is an IWA archive, store each archive using the file handler and, if
    specified, unpack the archives into the object handler."""
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
