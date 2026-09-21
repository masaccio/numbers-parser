import hashlib
import logging
import plistlib
import re
import struct
from abc import ABC, abstractmethod
from io import BytesIO
from os import urandom
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


class IWorkCrypto:
    def __init__(self, key: bytes):
        self._key = key

    @classmethod
    def from_password_verifier(cls, data: bytes, password: str | None):
        if len(data) != 104:
            msg = f"IWorkCrypto: unrecognized verifier format length ({len(data)} bytes)"
            raise ValueError(msg)

        # Unpack the 104-byte packed struct as little-endian
        version, format_version, iterations, salt, iv, encrypted_data = struct.unpack(
            "<HH I 16s 16s 64s",
            data,
        )

        if version != 2 or format_version != 1:
            msg = f"Unsupported version or format: {version}, {format_version}"
            raise ValueError(
                msg,
            )

        if not password:
            return None

        key = cls._create_key(password, salt, iterations)
        cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
        decryptor = cipher.decryptor()
        decrypted = decryptor.update(encrypted_data) + decryptor.finalize()

        # The last 32 bytes of the block should be equal to the SHA256 of the first 32 bytes
        hash_val = hashlib.sha256(decrypted[:32]).digest()
        if hash_val != decrypted[32:64]:
            return None

        return cls(key)

    @classmethod
    def from_password(cls, password: str):
        salt = urandom(16)
        iv = urandom(16)
        iterations = 100000
        key = cls._create_key(password, salt, iterations)

        verifier_payload = urandom(32)
        # The last 32 bytes of the block should be equal to the SHA256 of the first 32 bytes
        hash_val = hashlib.sha256(verifier_payload).digest()
        verifier_block = verifier_payload + hash_val

        cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
        encryptor = cipher.encryptor()
        encrypted_verifier_block = encryptor.update(verifier_block) + encryptor.finalize()

        verifier_data = struct.pack(
            "<HH I 16s 16s 64s",
            2,
            1,
            iterations,
            salt,
            iv,
            encrypted_verifier_block,
        )
        return verifier_data, cls(key)

    @staticmethod
    def _create_key(password: str, salt: bytes, iterations: int) -> bytes:
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA1(),  # noqa: S303
            length=16,
            salt=salt,
            iterations=iterations,
            backend=default_backend(),
        )
        return kdf.derive(password.encode("utf-8"))

    def decrypt_iwa(self, data: bytes) -> bytes:
        # The first 16 bytes are the IV and the last 20 bytes are garbage
        encrypted_length = len(data) - 36
        if encrypted_length < 16 or encrypted_length % 16:
            msg = "Data too short to be a valid encrypted IWA file"
            raise ValueError(msg)

        iv = data[:16]
        encrypted_bytes = data[16:-20]

        cipher = Cipher(algorithms.AES(self._key), modes.CBC(iv), backend=default_backend())
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

    def encrypt_iwa(self, data: bytes) -> bytes:
        iv = urandom(16)
        # 16 random bytes are placed at the head of the unencrypted payload
        header = urandom(16)
        # 20 bytes of garbage appended at the tail after encryption
        garbage = urandom(20)

        padder = padding.PKCS7(128).padder()
        padded_data = padder.update(header + data) + padder.finalize()

        cipher = Cipher(algorithms.AES(self._key), modes.CBC(iv), backend=default_backend())
        encryptor = cipher.encryptor()
        encrypted_bytes = encryptor.update(padded_data) + encryptor.finalize()

        return iv + encrypted_bytes + garbage


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
            properties_plist = self._zipf.read(max(metadata))

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
        self.is_encrypted = False

        doc_version = self.document_version
        if not self._handler.allowed_version(doc_version):
            warn(f"unsupported version '{doc_version}'", RuntimeWarning, stacklevel=2)

        if filepath.is_dir():
            self._read_package_encryption(self._filepath)
            self._read_objects_from_package(self._filepath)
        else:
            self._read_objects_from_zipfile(self._zipf)

    def save(
        self,
        filepath: Path,
        file_store: dict[str, object],
        package: bool = False,
        password: str | None = None,
        hint: str = "No hint",
    ) -> None:
        file_store = file_store.copy()
        crypto = None
        if password is not None:
            verifier_data, crypto = IWorkCrypto.from_password(password)
            file_store[".iwpv2"] = verifier_data
            file_store[".iwph"] = hint.encode("utf-8")

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
                    if crypto:
                        zipf.writestr(blob_path, crypto.encrypt_iwa(blob.to_buffer()))
                    else:
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
                    if crypto:
                        zipf.writestr(filepath_in_zip, crypto.encrypt_iwa(blob.to_buffer()))
                    else:
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
            elif sub_filepath.name in (".iwph", ".iwpv2"):
                continue
            else:
                with sub_filepath.open(mode="rb") as fh:
                    blob = fh.read()
                    package_filename = re.sub(r".*\.numbers/*", "", str(sub_filepath))
                    self._store_blob(package_filename, blob)

    def _read_package_encryption(self, filepath: Path) -> None:
        hint_filename = filepath / ".iwph"
        verifier_filename = filepath / ".iwpv2"
        if not hint_filename.exists() and not verifier_filename.exists():
            return
        if not hint_filename.is_file() or not verifier_filename.is_file():
            msg = "invalid Numbers document (missing encryption files)"
            raise FileError(msg)
        self._initialize_encryption(
            hint_filename.read_bytes().decode(),
            verifier_filename.read_bytes(),
        )

    def _initialize_encryption(self, hint: str, verifier_data: bytes) -> None:
        self.is_encrypted = True
        try:
            self._crypto = IWorkCrypto.from_password_verifier(verifier_data, self._password)
        except ValueError as e:
            msg = "Error initializing encryption verifier"
            raise FileError(msg) from e

        if not self._crypto:
            msg = f"Invalid password. Hint is '{hint}'"
            raise FileError(msg)

    def _read_objects_from_zipfile(self, zipf) -> None:
        try:
            hint = zipf.read(".iwph").decode()
        except KeyError:
            pass
        else:
            try:
                verifier_data = zipf.read(".iwpv2")
            except KeyError as e:
                msg = "invalid Numbers document (missing encryption files)"
                raise FileError(msg) from e
            self._initialize_encryption(hint, verifier_data)

        for filename in zipf.namelist():
            if filename in (".iwph", ".iwpv2"):
                continue
            blob = zipf.read(filename)
            if self.is_encrypted and filename.endswith(".iwa"):
                blob = self._crypto.decrypt_iwa(blob)

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
