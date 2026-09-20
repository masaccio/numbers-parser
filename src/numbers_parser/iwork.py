import hashlib
import logging
import plistlib
import re
import struct
from abc import ABC, abstractmethod
from io import BytesIO
from pathlib import Path
from sys import version_info
from warnings import warn
from zipfile import BadZipFile, ZipFile

from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes, padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

from numbers_parser.exceptions import FileError, FileFormatError
from numbers_parser.iwafile import IWAFile, is_iwa_file

logger = logging.getLogger(__name__)
debug = logger.debug


class IWorkHandler(ABC):
    @abstractmethod
    def store_file(self, filename: str, blob: bytes) -> None:
        """Store a protobuf archive."""

    @abstractmethod
    def store_object(self, filename: str, identifier: int, archive: object) -> None:
        """Store a binary blob of data from the iWork package."""

    @abstractmethod
    def allowed_format(self, extension: str) -> bool:
        """bool: Return ``True`` if the filename extension is supported by the handler."""

    @abstractmethod
    def allowed_version(self, version: str) -> bool:
        """bool: Return ``True`` if the document version is allowed."""


class IWPasswordVerifier:
    def __init__(self, data: bytes):
        if len(data) != 104:
            raise ValueError(f"IWPasswordVerifier: unrecognized format length ({len(data)} bytes)")

        # Unpack the 104-byte packed struct as little-endian[cite: 4]
        # uint16_t version, uint16_t format, uint32_t iterations, uint8_t salt[16], uint8_t iv[16], uint8_t data[64][cite: 4]
        self.version, self.format, self.iterations, self.salt, self.iv, self.data = struct.unpack(
            "<HH I 16s 16s 64s",
            data,
        )

        if self.version != 2 or self.format != 1:
            raise ValueError(
                f"Unsupported version or format: {self.version}, {self.format}[cite: 4]",
            )

    def create_key_with_password(self, password: str) -> bytes:
        if not password:
            return None

        # The 16-byte key is created using standard PBKDF2+SHA1[cite: 4]
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA1(),
            length=16,
            salt=self.salt,
            iterations=self.iterations,
            backend=default_backend(),
        )
        key = kdf.derive(password.encode("utf-8"))

        # Use the key to decrypt the 64-byte block[cite: 4]
        cipher = Cipher(algorithms.AES(key), modes.CBC(self.iv), backend=default_backend())
        decryptor = cipher.decryptor()

        decrypted = decryptor.update(self.data) + decryptor.finalize()

        # The last 32 bytes of the block should be equal to the SHA256 of the first 32 bytes[cite: 4]
        hash_val = hashlib.sha256(decrypted[:32]).digest()
        if hash_val != decrypted[32:64]:
            return None

        return key


class IWork:
    def __init__(self, handler: IWorkHandler = None) -> None:
        """
        Create an IWork document handler that can read and write iWork documents.

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
                msg = "invalid Numbers document (missing files)"
                raise FileFormatError(msg) from None
            with open(properties_filename, "rb") as fh:
                properties_plist = fh.read()
        else:
            metadata = [
                x.filename
                for x in self._zipf.filelist
                if x.filename.endswith(
                    ("Metadata/Properties.plist", "Metadata/BuildVersionHistory.plist"),
                )
            ]
            if len(metadata) != 2:
                msg = "invalid Numbers document (missing files)"
                raise FileFormatError(msg) from None
            properties_plist = self._zipf.read(sorted(metadata)[-1])

        try:
            doc_properties = plistlib.loads(properties_plist)
            doc_version = doc_properties["fileFormatVersion"]
        except plistlib.InvalidFileException:
            # Numbers allows malformed Properties.plist but not missing files
            doc_version = ""
            warn("can't read Numbers version from document", RuntimeWarning, stacklevel=2)
        return doc_version

    def open(self, filepath: Path, password: str | None) -> None:
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
        self._password = password
        if not filepath.exists():
            msg = "no such file or directory"
            raise FileError(msg)
        if not self._handler.allowed_format(filepath.suffix):
            msg = "invalid Numbers document (not a .numbers package/file)"
            raise FileFormatError(msg)

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

    def save(self, filepath: Path, file_store: dict[str, object], package: bool) -> None:
        if package:
            if filepath.is_dir():
                if filepath.suffix != ".numbers":
                    msg = "invalid Numbers document (not a Numbers package)"
                    raise FileFormatError(msg)
                if not (filepath / "Index.zip").is_file():
                    msg = "folder is not a numbers package"
                    raise FileFormatError(msg)
                # Test existing document is valid
                self._is_package = True
                _ = self.document_version
            elif filepath.is_file():
                msg = "cannot overwrite Numbers document file with package"
                raise FileFormatError(msg)
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

            for filepath_in_zip, blob in file_store.items():
                if isinstance(blob, IWAFile):
                    zipf.writestr(filepath_in_zip, blob.to_buffer())
                else:
                    zipf.writestr(filepath_in_zip, blob)
            zipf.close()

    def _open_zipfile(self, filepath: Path):
        """Open Zip file with the correct filename encoding supported by current python."""
        # Coverage is python version dependent, so one path with always fail coverage
        try:  # pragma: no cover
            if version_info >= (3, 11):
                return ZipFile(filepath, metadata_encoding="utf-8")
            return ZipFile(filepath)

        except BadZipFile:
            msg = "invalid Numbers document"
            raise FileFormatError(msg) from None

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
        try:
            _ = zipf.getinfo(".iwph")
        except KeyError:
            self.is_encrypted = False
        else:
            self.is_encrypted = True
            hint = zipf.read(".iwph").decode()
            verifier_data = zipf.read(".iwpv2")
            try:
                verifier = IWPasswordVerifier(verifier_data)
            except ValueError as e:
                msg = "Error initializing encryption verifier"
                raise FileError(msg) from e

            self._key = verifier.create_key_with_password(self._password)
            if not self._key:
                msg = f"Invalid password. Hint is '{hint}'"
                raise FileError(msg)

        for filename in zipf.namelist():
            blob = zipf.read(filename)
            if self.is_encrypted and filename.endswith(".iwa"):
                blob = self._decrypt_using_iwa_key(blob, self._key)

            if filename.lower().endswith("index.zip"):
                index_data = BytesIO(blob)
                self._read_objects_from_zipfile(self._open_zipfile(index_data))
            else:
                self._store_blob(filename, blob)

    def _store_blob(self, filename: str, blob: bytes) -> None:
        """
        If blob is an IWA archive, store each archive using the file handler and, if
        specified, unpack the archives into the object handler.
        """
        if filename.endswith(".iwa") and is_iwa_file(blob):
            try:
                iwaf = IWAFile.from_buffer(blob, filename)
            except Exception as e:
                msg = f"{filename}: invalid IWA file {filename}"
                raise FileFormatError(msg) from e

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

    def _decrypt_using_iwa_key(self, data: bytes, key: bytes) -> bytes:
        # The first 16 bytes are the IV and the last 20 bytes are garbage
        if len(data) < 36:
            msg = "Data too short to be a valid encrypted IWA file"
            raise ValueError(msg)

        iv = data[:16]
        encrypted_bytes = data[16:-20]

        cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
        decryptor = cipher.decryptor()

        decrypted_padded = decryptor.update(encrypted_bytes) + decryptor.finalize()

        unpadder = padding.PKCS7(128).unpadder()
        try:
            decrypted = unpadder.update(decrypted_padded) + unpadder.finalize()
        except ValueError as e:
            msg = "PKCS7 unpadding failed. Key might be correct but payload is corrupted."
            raise ValueError(msg) from e

        if len(decrypted) < 16:
            msg = "Decrypted data is too short to discard the 16-byte header"
            raise ValueError(msg)

        return decrypted[16:]
