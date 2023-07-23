from uuid import UUID, uuid1

from numbers_parser.exceptions import UnsupportedError
from numbers_parser.generated import TSPMessages_pb2 as TSPMessages


class NumbersUUID(UUID):
    def __init__(self, uuid=None):
        if uuid is None:
            super().__init__(int=uuid1().int)
        elif type(uuid) == int:
            super().__init__(int=uuid)
        elif type(uuid) == str:
            super().__init__(hex=uuid)
        elif type(uuid) == TSPMessages.UUID:
            uuid_int = uuid.upper << 64 | uuid.lower
            super().__init__(int=uuid_int)
        elif type(uuid) == TSPMessages.CFUUIDArchive:
            uuid_int = (
                (uuid.uuid_w3 << 96) | (uuid.uuid_w2 << 64) | (uuid.uuid_w1 << 32) | uuid.uuid_w0
            )
            super().__init__(int=uuid_int)
        elif type(uuid) == dict:
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
                raise UnsupportedError("Unsupported UUID dict structure")
        else:  # pragma: no cover
            raise UnsupportedError(f"Unsupported UUID init type {type(uuid).__name__}")

    @property
    def dict2(self) -> dict:
        upper = self.int >> 64
        lower = self.int & 0xFFFFFFFFFFFFFFFF
        return {"upper": upper, "lower": lower}

    @property
    def dict4(self) -> object:
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
    def protobuf2(self) -> object:
        upper = self.int >> 64
        lower = self.int & 0xFFFFFFFFFFFFFFFF
        return TSPMessages.UUID(upper=upper, lower=lower)

    @property
    def protobuf4(self) -> object:
        uuid_w3 = self.int >> 96
        uuid_w2 = (self.int >> 64) & 0xFFFFFFFF
        uuid_w1 = (self.int >> 32) & 0xFFFFFFFF
        uuid_w0 = self.int & 0xFFFFFFFF
        return TSPMessages.CFUUIDArchive(
            uuid_w3=uuid_w3, uuid_w2=uuid_w2, uuid_w1=uuid_w1, uuid_w0=uuid_w0
        )
