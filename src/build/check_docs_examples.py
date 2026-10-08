"""
Check the protobuf snippets in docs/api/file-format.rst against the real
protobuf descriptors and against a document loaded by numbers-parser.

For every ``.. code-block:: protobuf`` block in the RST file this script:

1. extracts the snippet text,
2. compiles it (a small proto2 parser producing message/enum/field specs),
3. resolves each message to a compiled descriptor in numbers_parser.generated
   and verifies every documented field (name, number, label, type, nesting),
4. loads src/numbers_parser/data/empty.numbers, walks doc._model.objects (and
   the IWA ArchiveInfo/MessageInfo headers) and checks that live instances
   reflect the documented functionality: required fields set, and
   TSP.Reference fields resolving to objects in the store.

Usage: python src/build/check_docs_examples.py [--rst PATH] [--doc PATH] [-v]
Exit status is non-zero if any check fails.
"""

from __future__ import annotations

import argparse
import importlib
import pkgutil
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from google.protobuf.descriptor import FieldDescriptor as FD  # noqa: N817

from numbers_parser import Document, generated

SCALARS = {
    "double": FD.TYPE_DOUBLE,
    "float": FD.TYPE_FLOAT,
    "int32": FD.TYPE_INT32,
    "int64": FD.TYPE_INT64,
    "uint32": FD.TYPE_UINT32,
    "uint64": FD.TYPE_UINT64,
    "sint32": FD.TYPE_SINT32,
    "sint64": FD.TYPE_SINT64,
    "fixed32": FD.TYPE_FIXED32,
    "fixed64": FD.TYPE_FIXED64,
    "sfixed32": FD.TYPE_SFIXED32,
    "sfixed64": FD.TYPE_SFIXED64,
    "bool": FD.TYPE_BOOL,
    "string": FD.TYPE_STRING,
    "bytes": FD.TYPE_BYTES,
}
LABELS = {"optional", "required", "repeated"}


def is_repeated(fd) -> bool:
    """Protobuf >= 6 removed FieldDescriptor.label; support both APIs."""
    if hasattr(fd, "is_repeated"):
        return bool(fd.is_repeated)
    return fd.label == FD.LABEL_REPEATED


def is_required(fd) -> bool:
    if hasattr(fd, "is_required"):
        return bool(fd.is_required)
    return fd.label == FD.LABEL_REQUIRED


def label_of(fd) -> str:
    if is_repeated(fd):
        return "repeated"
    if is_required(fd):
        return "required"
    return "optional"


def is_packed(fd) -> bool:
    if hasattr(fd, "is_packed"):
        return bool(fd.is_packed)
    opts = fd.GetOptions()
    return bool(opts.packed) if opts.HasField("packed") else False


# --------------------------------------------------------------------------
# Snippet extraction and compilation
# --------------------------------------------------------------------------
@dataclass
class Field:
    label: str
    type: str
    name: str
    number: int
    options: str = ""


@dataclass
class Message:
    name: str
    fields: list[Field] = field(default_factory=list)
    messages: list[Message] = field(default_factory=list)
    enums: dict[str, dict[str, int]] = field(default_factory=dict)


@dataclass
class Snippet:
    index: int
    line: int
    text: str
    messages: list[Message] = field(default_factory=list)


def extract_snippets(rst_path: Path) -> list[Snippet]:
    lines = rst_path.read_text(encoding="utf-8").splitlines()
    snippets = []
    i = 0
    while i < len(lines):
        if re.match(r"^\.\. code-block::\s+protobuf\s*$", lines[i]):
            start = i + 1
            i += 1
            body = []
            while i < len(lines) and (lines[i].strip() == "" or lines[i].startswith(" ")):
                body.append(lines[i])
                i += 1
            indent = min((len(x) - len(x.lstrip()) for x in body if x.strip()), default=0)
            text = "\n".join(x[indent:] for x in body).strip("\n")
            snippets.append(Snippet(len(snippets) + 1, start, text))
        else:
            i += 1
    return snippets


TOKEN_RE = re.compile(
    r"\s*(//[^\n]*|[A-Za-z_][\w.]*|\.[A-Za-z_][\w.]*|-?\d+|0x[0-9a-fA-F]+|[{}=;\[\],])"
)


def tokenize(text: str) -> list[str]:
    tokens = []
    pos = 0
    text = text.rstrip()
    while pos < len(text):
        m = TOKEN_RE.match(text, pos)
        if not m:
            if text[pos:].strip() == "":
                break
            msg = f"cannot tokenize near {text[pos : pos + 30]!r}"
            raise SyntaxError(msg)
        pos = m.end()
        tok = m.group(1)
        if tok.startswith("//"):
            continue
        tokens.append(tok)
    return tokens


def compile_snippet(snippet: Snippet) -> None:
    """Parse the protobuf snippet into Message specs."""
    toks = tokenize(snippet.text)
    pos = 0

    def expect(value=None):
        nonlocal pos
        if pos >= len(toks) or (value is not None and toks[pos] != value):
            got = toks[pos] if pos < len(toks) else "EOF"
            msg = f"expected {value!r}, got {got!r}"
            raise SyntaxError(msg)
        pos += 1
        return toks[pos - 1]

    def parse_enum():
        name = expect()
        expect("{")
        values = {}
        while toks[pos] != "}":
            vname = expect()
            expect("=")
            values[vname] = int(expect(), 0)
            expect(";")
        expect("}")
        return name, values

    def parse_message():
        msg = Message(expect())
        expect("{")
        while toks[pos] != "}":
            tok = toks[pos]
            if tok == "message":
                expect()
                msg.messages.append(parse_message())
            elif tok == "enum":
                expect()
                n, v = parse_enum()
                msg.enums[n] = v
            else:
                label = expect()
                if label not in LABELS:
                    msg_ = f"unknown label {label!r} in message {msg.name}"
                    raise SyntaxError(msg_)
                ftype = expect()
                fname = expect()
                expect("=")
                num = int(expect(), 0)
                options = ""
                if toks[pos] == "[":
                    expect()
                    while toks[pos] != "]":
                        options += expect()
                    expect("]")
                expect(";")
                msg.fields.append(Field(label, ftype, fname, num, options))
        expect("}")
        return msg

    while pos < len(toks):
        expect("message")
        snippet.messages.append(parse_message())


# --------------------------------------------------------------------------
# Descriptor registry
# --------------------------------------------------------------------------
def load_generated():
    """Import every generated *_pb2 module; return {full_name: descriptor}."""
    registry = {}

    def walk(desc):
        registry[desc.full_name] = desc
        for nested in desc.nested_types:
            walk(nested)

    for mod in pkgutil.iter_modules(generated.__path__):
        if not mod.name.endswith("_pb2"):
            continue
        m = importlib.import_module(f"numbers_parser.generated.{mod.name}")
        for desc in m.DESCRIPTOR.message_types_by_name.values():
            walk(desc)
    return registry


def check_field(desc, f: Field, errors: list[str]) -> None:
    fd = desc.fields_by_name.get(f.name)
    if fd is None:
        errors.append(f"{desc.full_name}: no field {f.name!r}")
        return
    if fd.number != f.number:
        errors.append(f"{desc.full_name}.{f.name}: number {fd.number} != documented {f.number}")
    actual_label = label_of(fd)
    if actual_label != f.label:
        errors.append(f"{desc.full_name}.{f.name}: label {actual_label} != documented {f.label}")
    if f.type in SCALARS:
        if fd.type != SCALARS[f.type]:
            errors.append(f"{desc.full_name}.{f.name}: type is not {f.type}")
    else:
        target = fd.message_type or fd.enum_type
        if target is None:
            errors.append(f"{desc.full_name}.{f.name}: documented {f.type} but field is scalar")
        else:
            doc_name = f.type.lstrip(".")
            if (f.type.startswith(".") and target.full_name != doc_name) or (
                not f.type.startswith(".") and target.name != doc_name.split(".")[-1]
            ):
                errors.append(f"{desc.full_name}.{f.name}: type {target.full_name} != {doc_name}")
    if "packed=true" in f.options and not is_packed(fd):
        errors.append(f"{desc.full_name}.{f.name}: documented packed but field is not packed")


def check_message(desc, spec: Message, errors: list[str]) -> None:
    for f in spec.fields:
        check_field(desc, f, errors)
    for ename, values in spec.enums.items():
        ed = desc.enum_types_by_name.get(ename)
        if ed is None:
            errors.append(f"{desc.full_name}: no enum {ename!r}")
            continue
        for vname, vnum in values.items():
            ev = ed.values_by_name.get(vname)
            if ev is None or ev.number != vnum:
                errors.append(f"{ed.full_name}: value {vname}={vnum} mismatch")
    for nested in spec.messages:
        nd = desc.nested_types_by_name.get(nested.name)
        if nd is None:
            errors.append(f"{desc.full_name}: no nested message {nested.name!r}")
        else:
            check_message(nd, nested, errors)


def match_descriptor(registry, spec: Message):
    """Return (descriptor, errors) for the best-matching candidate of a top-level spec."""
    cands = [d for d in registry.values() if d.containing_type is None and d.name == spec.name]
    best = None
    for d in cands:
        errs: list[str] = []
        check_message(d, spec, errs)
        if not errs:
            return d, []
        if best is None or len(errs) < len(best[1]):
            best = (d, errs)
    if best is None:
        return None, [f"no generated message named {spec.name!r}"]
    return best


# --------------------------------------------------------------------------
# Live document checks
# --------------------------------------------------------------------------
def iter_messages(msg):
    """Yield msg and all embedded sub-messages."""
    yield msg
    for fdesc, value in msg.ListFields():
        if fdesc.type != FD.TYPE_MESSAGE:
            continue
        for item in value if is_repeated(fdesc) else [value]:
            # map fields yield keys; only descend into real messages
            if hasattr(item, "ListFields"):
                yield from iter_messages(item)


def collect_instances(doc) -> dict[str, list]:
    """Index every message (including embedded and IWA header messages) by full name."""
    store = doc._model.objects
    instances: dict[str, list] = {}
    for obj_id in store._objects:
        for m in iter_messages(store[obj_id]):
            instances.setdefault(m.DESCRIPTOR.full_name, []).append(m)
    for iwa in store.file_store.values():
        for chunk in getattr(iwa, "chunks", []):
            for archive in chunk.archives:
                for m in iter_messages(archive.header):
                    instances.setdefault(m.DESCRIPTOR.full_name, []).append(m)
    return instances


def check_live(desc, spec: Message, instances, store, errors: list[str]) -> int:
    """Check documented functionality on live instances; return count checked."""
    live = instances.get(desc.full_name, [])
    for inst in live:
        for f in spec.fields:
            fd = desc.fields_by_name.get(f.name)
            if fd is None:
                continue
            if f.label == "required" and not inst.HasField(f.name):
                errors.append(f"{desc.full_name}: required field {f.name} unset")
            if fd.type == FD.TYPE_MESSAGE and fd.message_type.full_name == "TSP.Reference":
                if is_repeated(fd):
                    refs = list(getattr(inst, f.name))
                else:
                    refs = [getattr(inst, f.name)] if inst.HasField(f.name) else []
                for ref in refs:
                    if ref.identifier not in store:
                        errors.append(
                            f"{desc.full_name}.{f.name}: reference {ref.identifier} "
                            "not found in doc._model.objects"
                        )
        if desc.full_name == "TSP.ArchiveInfo":
            total = sum(mi.length for mi in inst.message_infos)
            if total < 0:
                errors.append("ArchiveInfo: negative payload length")
    for nested in spec.messages:
        nd = desc.nested_types_by_name.get(nested.name)
        if nd is not None:
            check_live(nd, nested, instances, store, errors)
    return len(live)


# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--rst", type=Path, default=Path("docs/api/file-format.rst"))
    ap.add_argument("--doc", type=Path, default=Path("src/numbers_parser/data/empty.numbers"))
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    snippets = extract_snippets(args.rst)
    if not snippets:
        print(f"No protobuf snippets found in {args.rst}")
        return 1

    registry = load_generated()
    doc = Document(args.doc)
    store = doc._model.objects
    instances = collect_instances(doc)

    failures = 0
    for snip_num, snip in enumerate(snippets, start=1):
        if args.verbose:
            print(f"Snippet #{snip_num}:\n------ BEGINS ------\n{snip.text}\n------ ENDS ------")
        try:
            compile_snippet(snip)
        except SyntaxError as e:
            print(f"FAIL snippet {snip.index} (line {snip.line}): compile error: {e}")
            failures += 1
            continue
        for spec in snip.messages:
            where = f"snippet {snip.index} (line {snip.line}) {spec.name}"
            desc, errors = match_descriptor(registry, spec)
            if desc is None or errors:
                failures += 1
                print(f"FAIL {where}")
                for e in errors:
                    print(f"     {e}")
                continue
            live_errors: list[str] = []
            n = check_live(desc, spec, instances, store, live_errors)
            if live_errors:
                failures += 1
                print(f"FAIL {where} -> {desc.full_name}: live document mismatch")
                for e in sorted(set(live_errors))[:10]:
                    print(f"     {e}")
            else:
                status = f"{n} live instance(s)" if n else "schema only, no instance in doc"
                if args.verbose:
                    print(f"PASS {where} -> {desc.full_name} ({status})")

    print(f"check_docs_examples: {len(snippets)} snippets checked, {failures} failure(s)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
