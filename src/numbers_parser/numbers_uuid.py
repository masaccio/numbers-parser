"""Numbers uuid module helpers."""
from uuid import UUID, uuid1

from numbers_parser.exceptions import UnsupportedError
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


class NumbersUUID(UUID):
    """Convert UUIDs between archive representations."""

    def __init__(
        self,
        uuid: int
        | str
        | TSPMessages.UUID
        | TSPMessages.CFUUIDArchive
        | dict[str, int | str]
        | None = None,
    ) -> None:
        """Initialize the instance state."""
        if uuid is None:
            super().__init__(int=uuid1().int)
        elif isinstance(uuid, int):
            super().__init__(int=uuid)
        elif isinstance(uuid, str):
            super().__init__(hex=uuid)
        elif isinstance(uuid, TSPMessages.UUID):
            uuid_int = uuid.upper << 64 | uuid.lower
            super().__init__(int=uuid_int)
        elif isinstance(uuid, TSPMessages.CFUUIDArchive):
            uuid_int = (
                (uuid.uuid_w3 << 96) | (uuid.uuid_w2 << 64) | (uuid.uuid_w1 << 32) | uuid.uuid_w0
            )
            super().__init__(int=uuid_int)
        elif isinstance(uuid, dict):
            if "uuid_w0" in uuid and "uuid_w1" in uuid:
                uuid_int = (
                    (int(uuid["uuid_w3"]) << 96)
                    | (int(uuid["uuid_w2"]) << 64)
                    | (int(uuid["uuid_w1"]) << 32)
                    | int(uuid["uuid_w0"])
                )
                super().__init__(int=uuid_int)
            elif "upper" in uuid and "lower" in uuid:
                uuid_int = int(uuid["upper"]) << 64 | int(uuid["lower"])
                super().__init__(int=uuid_int)
            else:
                msg = "Unsupported UUID dict structure"
                raise UnsupportedError(msg)
        else:
            msg = f"Unsupported UUID init type {type(uuid).__name__}"
            raise UnsupportedError(msg)

    @property
    def dict2(self) -> dict:
        """Create a UUID from the two-part dictionary representation."""
        upper = self.int >> 64
        lower = self.int & 0xFFFFFFFFFFFFFFFF
        return {"upper": upper, "lower": lower}

    @property
    def dict4(self) -> dict[str, int]:
        """Create a UUID from the four-part dictionary representation."""
        uuid_w3 = self.int >> 96
        uuid_w2 = (self.int >> 64) & 0xFFFFFFFF
        uuid_w1 = (self.int >> 32) & 0xFFFFFFFF
        uuid_w0 = self.int & 0xFFFFFFFF
        return {
            "uuid_w3": uuid_w3,
            "uuid_w2": uuid_w2,
            "uuid_w1": uuid_w1,
            "uuid_w0": uuid_w0,
        }

    @property
    def protobuf2(self) -> TSPMessages.UUID:
        """Create the two-part protobuf representation of the UUID."""
        upper = self.int >> 64
        lower = self.int & 0xFFFFFFFFFFFFFFFF
        return TSPMessages.UUID(upper=upper, lower=lower)

    @property
    def protobuf4(self) -> TSPMessages.CFUUIDArchive:
        """Create the four-part protobuf representation of the UUID."""
        uuid_w3 = self.int >> 96
        uuid_w2 = (self.int >> 64) & 0xFFFFFFFF
        uuid_w1 = (self.int >> 32) & 0xFFFFFFFF
        uuid_w0 = self.int & 0xFFFFFFFF
        return TSPMessages.CFUUIDArchive(
            uuid_w3=uuid_w3,
            uuid_w2=uuid_w2,
            uuid_w1=uuid_w1,
            uuid_w0=uuid_w0,
        )


def uuid_to_hex(archive: object) -> str:
    """Convert a protobuf UUID to a hex string."""
    uuid = NumbersUUID(archive)
    return uuid.hex
