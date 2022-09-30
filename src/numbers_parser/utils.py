from numbers_parser.exceptions import UnsupportedError


def uuid(ref: dict) -> int:
    """
    Extract storage buffers for each cell in a table row

    Args:
        ref: Google protobuf containing either four 32-bit IDs or two 64-bit IDs

    Returns:
        uuid: 128-bit UUID

    Raises:
        UnsupportedError: object does not include expected UUID fields
    """
    try:
        if hasattr(ref, "upper"):
            uuid = ref.upper << 64 | ref.lower
        else:
            uuid = (
                (ref.uuid_w3 << 96)
                | (ref.uuid_w2 << 64)
                | (ref.uuid_w1 << 32)
                | ref.uuid_w0
            )
    except AttributeError:
        raise UnsupportedError(f"Unsupported UUID structure: {ref}")  # pragma: no cover

    return uuid
