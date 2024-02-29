import logging
import plistlib
import re
from io import BytesIO
from pathlib import Path
from sys import version_info
from typing import Dict
from warnings import warn
from zipfile import BadZipFile, ZipFile

from numbers_parser.exceptions import FileError, FileFormatError
from numbers_parser.iwafile import IWAFile, is_iwa_file

logger = logging.getLogger(__name__)
debug = logger.debug


class IWorkHandler:
    def __init__(self):
        pass  # pragma: nocover

    def store_file(self, filename: str, blob: bytes) -> None:
        """Store a profobuf archive."""
        pass  # pragma: nocover

    def store_object(self, filename: str, identifier: int, archive: object) -> None:
        """Store a binary blob of data from the iWork package."""
        pass  # pragma: nocover

    def allowed_format(self, extension: str) -> bool:
        """bool: Return ``True`` if the filename extension is supported by the handler."""
        pass  # pragma: nocover

    def allowed_version(self, version: str) -> bool:
        """bool: Return ``True`` if the document version is allowed."""
        pass  # pragma: nocover


class IWork:
    def __init__(self, handler: IWorkHandler = None):
        """
        Create an IWork document handler that can read and write iWork documents

        Parameters
        ----------
        handler: IWorkHandler, optional
            The handler that is called to store objects and files and to check
            versions and supported document formats.
        """
        self._handler = handler

    @property
    def document_version(self) -> str:
        """
        str: the version of the iWork document.

        Raises
        ------
        FileFormatError:
            If document version cannot be read from the document.
        """
        if self._is_package:
            properties_filename = self._filepath / "Metadata/Properties.plist"
            build_filename = self._filepath / "Metadata/BuildVersionHistory.plist"
            if not properties_filename.exists() or not build_filename.exists():
                raise FileFormatError("invalid Numbers document (missing files)") from None
            with open(properties_filename, "rb") as fh:
                properties_plist = fh.read()
        else:
            metadata = [
                x.filename
                for x in self._zipf.filelist
                if x.filename.endswith(
                    ("Metadata/Properties.plist", "Metadata/BuildVersionHistory.plist")
                )
            ]
            if len(metadata) != 2:
                raise FileFormatError("invalid Numbers document (missing files)") from None
            properties_plist = self._zipf.read(sorted(metadata)[-1])

        try:
            doc_properties = plistlib.loads(properties_plist)
            doc_version = doc_properties["fileFormatVersion"]
        except plistlib.InvalidFileException:
            # Numbers allows malformed Properties.plist but not missing files
            doc_version = ""
            warn("can't read Numbers version from document", RuntimeWarning, stacklevel=2)
        return doc_version

    def open(self, filepath: Path) -> None:
        """
        Open an iWork file and read in the files and archives contained in it.

        Raises
        ------
        FileFormatError
            If any errors occur extracting data from the archive

        Warns
        -----
        RuntimeWarning
            If the version of the document is one that the IWork blob handler
            reports is unsupported.
        """
        debug("open: filename=%s", filepath)
        self._filepath = filepath
        if not filepath.exists():
            raise FileError("no such file or directory")
        if not self._handler.allowed_format(filepath.suffix):
            raise FileFormatError("invalid Numbers document (not a .numbers package/file)")

        if filepath.is_dir():
            self._is_package = True
        else:
            self._is_package = False
            self._zipf = self._open_zipfile(filepath)

        doc_version = self.document_version
        if not self._handler.allowed_version(doc_version):
            warn(f"unsupported version '{doc_version}'", RuntimeWarning, stacklevel=2)

        if filepath.is_dir():
            self._read_objects_from_package(self._filepath)
        else:
            self._read_objects_from_zipfile(self._zipf)

    def save(self, filepath: Path, file_store: Dict[str, object], package: bool):
        if package:
            if filepath.is_dir():
                if not filepath.suffix == ".numbers":
                    raise FileFormatError("invalid Numbers document (not a Numbers package)")
                if not (filepath / "Index.zip").is_file():
                    raise FileFormatError("folder is not a numbers package")
                # Test existing document is valid
                self._is_package = True
                _ = self.document_version
            elif filepath.is_file():
                raise FileFormatError("cannot overwrite Numbers document file with package")
            else:
                filepath.mkdir()

            # OSError possible exception; allow it to propagate up
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
            # OSError possible exception; allow it to propagate up
            zipf = ZipFile(filepath, "w")

            for filepath, blob in file_store.items():
                if isinstance(blob, IWAFile):
                    zipf.writestr(filepath, blob.to_buffer())
                else:
                    zipf.writestr(filepath, blob)
            zipf.close()

    def _open_zipfile(self, filepath: Path):
        """Open Zip file with the correct filename encoding supported by current python"""
        # Coverage is python version dependent, so one path with always fail coverage
        try:
            if version_info.minor >= 11:  # pragma: no cover
                return ZipFile(filepath, metadata_encoding="utf-8")
            else:  # pragma: no cover
                return ZipFile(filepath)
        except BadZipFile:
            raise FileFormatError("invalid Numbers document") from None

    def _read_objects_from_package(self, filepath: Path) -> None:
        """
        Read a Numbers package and iterate through all files and directories
        storing the files blobs and objects though the supplies callbacks.
        """
        for sub_filepath in filepath.iterdir():
            if sub_filepath.is_dir():
                self._read_objects_from_package(sub_filepath)
            elif sub_filepath.name.lower() == "index.zip":
                zipf = self._open_zipfile(sub_filepath)
                self._read_objects_from_zipfile(zipf)
            else:
                with sub_filepath.open(mode="rb") as fh:
                    blob = fh.read()
                    package_filename = re.sub(r".*\.numbers/*", "", str(sub_filepath))
                    self._store_blob(package_filename, blob)

    def _read_objects_from_zipfile(self, zipf) -> None:
        for filename in zipf.namelist():
            blob = zipf.read(filename)
            if filename.lower().endswith("index.zip"):
                index_data = BytesIO(blob)
                self._read_objects_from_zipfile(self._open_zipfile(index_data))
            else:
                self._store_blob(filename, blob)

    def _store_blob(self, filename: str, blob: bytes) -> None:
        """If blob is an IWA archive, store each archive using the file handler and, if
        specified, unpack the archives into the object handler."""
        if filename.endswith(".iwa") and is_iwa_file(blob):
            try:
                iwaf = IWAFile.from_buffer(blob, filename)
            except Exception as e:
                raise FileFormatError(f"{filename}: invalid IWA file {filename}") from e

            # Data from Numbers always has just one chunk. Some archives
            # have multiple objects though they appear not to contain
            # useful data.
            for archive in iwaf.chunks[0].archives:
                identifier = archive.header.identifier
                debug("store IWA: filename=%s", filename)
                self._handler.store_object(filename, identifier, archive.objects[0])

            self._handler.store_file(filename, iwaf)
        else:
            debug("store blob: filename=%s", filename)
            self._handler.store_file(filename, blob)
