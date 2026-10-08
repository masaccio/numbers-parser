"""
Compatibility helpers that give python-betterproto2 messages the parts of the
Google protobuf API that numbers-parser relies on.

* Reading an unset singular message field returns an empty message, and reading an
  unset optional scalar returns its default (like Google protobuf) rather than ``None``. The empty message is not stored, so presence
  is still reported correctly by :func:`has_field` and serialization is unchanged.
  Assigning to a field of an unset sub-message therefore has no effect: assign
  the sub-message instead.
* :func:`has_field`, :func:`list_fields`, :func:`merge_from`, :func:`copy_from`
  and :func:`clear_field` replace ``HasField``, ``ListFields``, ``MergeFrom``,
  ``CopyFrom`` and ``ClearField``.
"""

import dataclasses
import threading
from copy import deepcopy
from typing import Any

import betterproto2
from betterproto2 import TYPE_MAP, TYPE_MESSAGE, FieldMetadata, Message, casing

_state = threading.local()


def _vivify_enabled() -> bool:
    return getattr(_state, "depth", 0) == 0


_ZERO_DEFAULTS = {
    betterproto2.TYPE_BOOL: False,
    betterproto2.TYPE_FLOAT: 0.0,
    betterproto2.TYPE_DOUBLE: 0.0,
    betterproto2.TYPE_STRING: "",
    betterproto2.TYPE_BYTES: b"",
}


class _UnsetDefault:
    """Data descriptor returning the protobuf default for unset optional fields."""

    __slots__ = ("default", "is_message", "name")

    def __init__(self, name: str, default: Any, is_message: bool = False) -> None:
        self.name = name
        self.default = default
        self.is_message = is_message

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        try:
            value = obj.__dict__[self.name]
        except KeyError:
            return None
        if value is None and _vivify_enabled():
            return self.default(obj) if callable(self.default) else self.default
        return value

    def __set__(self, obj, value) -> None:
        if isinstance(value, dict) and self.is_message:
            value = type(obj)._betterproto.cls_by_field[self.name].from_dict(value)
        obj.__dict__[self.name] = value


def _message_default(name: str):
    return lambda obj: type(obj)._betterproto.cls_by_field[name]()


def _enum_default(name: str, number: int):
    return lambda obj: type(obj)._betterproto.cls_by_field[name](number)


def install_message_fields(
    namespace: dict[str, Any], defaults: dict[str, dict[str, Any]] | None = None
) -> None:
    """
    Install the protobuf read semantics on every message class in a generated module.

    Unset singular message fields read as an empty message and unset optional scalar
    fields read as their (proto2) default value. ``defaults`` maps class names to
    declared default values: plain values, or ``("enum", number)`` for enums.
    """
    defaults = defaults or {}
    for cls_name, cls in list(namespace.items()):
        if not (
            isinstance(cls, type)
            and issubclass(cls, Message)
            and cls.__module__ == namespace["__name__"]
        ):
            continue
        declared = defaults.get(cls_name, {})
        for field in dataclasses.fields(cls):
            meta = FieldMetadata.get(field)
            if meta.repeated or meta.proto_type == TYPE_MAP:
                continue
            is_message = meta.proto_type == TYPE_MESSAGE
            if is_message:
                default = _message_default(field.name)
            elif not meta.optional:
                continue
            elif field.name in declared:
                value = declared[field.name]
                default = _enum_default(field.name, value[1]) if isinstance(value, tuple) else value
            elif meta.proto_type == betterproto2.TYPE_ENUM:
                default = _enum_default(field.name, 0)
            else:
                default = _ZERO_DEFAULTS.get(meta.proto_type, 0)
            setattr(cls, field.name, _UnsetDefault(field.name, default, is_message))


# Field names are kept exactly as defined in the protobuf files, so dictionaries
# and JSON use the original names rather than snake_case.
betterproto2.safe_snake_case = casing.sanitize_name


def identity_casing(name: str) -> str:
    return name


def _plain_reads(func):
    def wrapper(*args, **kwargs):
        _state.depth = getattr(_state, "depth", 0) + 1
        try:
            return func(*args, **kwargs)
        finally:
            _state.depth -= 1

    wrapper.__name__ = func.__name__
    wrapper.__doc__ = func.__doc__
    return wrapper


for _name in (
    "__bytes__",
    "__eq__",
    "__repr__",
    "__bool__",
    "__deepcopy__",
    "to_dict",
    "is_set",
    "load",
):
    setattr(Message, _name, _plain_reads(getattr(Message, _name)))


def _field_default(msg: Message, name: str) -> Any:
    return msg._betterproto.default_gen[name]()


def has_field(msg: Message, name: str) -> bool:
    """Equivalent of protobuf ``HasField`` (also true for non-empty repeated fields)."""
    value = msg.__dict__.get(name)
    if value is None:
        return False
    return not (isinstance(value, (list, dict)) and len(value) == 0)


def list_fields(msg: Message) -> list[tuple[str, FieldMetadata, Any]]:
    """Equivalent of protobuf ``ListFields``: the populated fields in field-number order."""
    meta_by_name = msg._betterproto.meta_by_field_name
    result = []
    for name in msg._betterproto.sorted_field_names:
        if has_field(msg, name):
            result.append((name, meta_by_name[name], msg.__dict__[name]))
    return result


def clear_field(msg: Message, name: str) -> None:
    """Equivalent of protobuf ``ClearField``."""
    setattr(msg, name, _field_default(msg, name))


def merge_from(dst: Message, src: Message) -> None:
    """Equivalent of protobuf ``MergeFrom``."""
    if dst is src:
        return
    for name, meta, value in list_fields(src):
        if meta.repeated:
            getattr(dst, name).extend(deepcopy(value))
        elif meta.proto_type == TYPE_MAP:
            getattr(dst, name).update(deepcopy(value))
        elif meta.proto_type == TYPE_MESSAGE:
            if has_field(dst, name):
                merge_from(dst.__dict__[name], value)
            else:
                setattr(dst, name, deepcopy(value))
        else:
            setattr(dst, name, value)
    dst._unknown_fields += src._unknown_fields


def mutable(msg: Message, name: str) -> Message:
    """Return a sub-message of ``msg``, creating and storing it first if it is unset."""
    value = msg.__dict__.get(name)
    if value is None:
        value = type(msg)._betterproto.cls_by_field[name]()
        setattr(msg, name, value)
    return value


def set_path(msg: Message, path: str, value: Any) -> None:
    """Set a (possibly nested) field such as ``a.b.c`` creating any unset sub-messages."""
    *parents, leaf = path.split(".")
    for name in parents:
        msg = mutable(msg, name)
    setattr(msg, leaf, value)


def merge_field(msg: Message, name: str, src: Message) -> None:
    """Equivalent of protobuf ``msg.name.MergeFrom(src)`` where ``name`` is a message field."""
    merge_from(mutable(msg, name), src)


def copy_from(dst: Message, src: Message) -> None:
    """Equivalent of protobuf ``CopyFrom``."""
    if dst is src:
        return
    for name in dst._betterproto.sorted_field_names:
        clear_field(dst, name)
    dst._unknown_fields = b""
    merge_from(dst, src)


__all__ = [
    "betterproto2",
    "clear_field",
    "copy_from",
    "has_field",
    "identity_casing",
    "install_message_fields",
    "list_fields",
    "merge_field",
    "merge_from",
    "mutable",
    "set_path",
]
