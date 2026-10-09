"""
Check the Python code examples in the numbers-parser sources and docs.

Scans the hand-written modules in src/numbers_parser (not ``generated``) and
the RST files in docs/ for ``.. code:: python`` / ``.. code-block:: python``
blocks. Two kinds of block are recognised:

* interpreter sessions (lines prefixed ``>>>``) whose output is compared with
  the documented output, ignoring memory addresses, and
* full code, which is simply executed.

Document.__init__ is monkey patched so any filename opens
tests/data/check-docs-examples.numbers; Document.save is redirected to a
temporary directory.

Usage: python src/build/check_docs_examples.py [-v]
Exit status is non-zero if any example fails.
"""

from __future__ import annotations

import argparse
import ast
import contextlib
import doctest
import io
import os
import re
import sys
import tempfile
import textwrap
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numbers_parser
from numbers_parser import *  # noqa: F403
from numbers_parser import Document as RealDocument

ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "tests" / "data"
DATA_FILE = DATA_DIR / "check-docs-examples.numbers"
DIRECTIVE_RE = re.compile(r"^(\s*)\.\. code(?:-block)?::\s+python\s*$")
ADDRESS_RE = re.compile(r"0x[0-9a-fA-F]+")


@dataclass
class Example:
    path: Path
    line: int
    text: str

    @property
    def is_session(self) -> bool:
        return any(x.lstrip().startswith(">>>") for x in self.text.splitlines())

    @property
    def where(self) -> str:
        return f"{self.path.relative_to(ROOT)}:{self.line}"


def extract_examples(path: Path) -> list[Example]:
    lines = path.read_text(encoding="utf-8").splitlines()
    examples = []
    i = 0
    while i < len(lines):
        m = DIRECTIVE_RE.match(lines[i])
        if not m:
            i += 1
            continue
        indent = len(m.group(1))
        start = i + 1
        i += 1
        body = []
        while i < len(lines):
            line = lines[i]
            if line.strip() and len(line) - len(line.lstrip()) <= indent:
                break
            body.append(line)
            i += 1
        text = textwrap.dedent("\n".join(body)).strip("\n")
        if text:
            examples.append(Example(path, start + 1, text))
    return examples


def find_examples() -> list[Example]:
    files = sorted(p for p in (ROOT / "src" / "numbers_parser").glob("*.py"))
    files += sorted((ROOT / "docs").rglob("*.rst"))
    examples = []
    for f in files:
        examples.extend(extract_examples(f))
    return examples


def make_document_shim(tmpdir: Path):
    original_save = RealDocument.save

    def save(self, filename, *args, **kwargs):
        return original_save(self, tmpdir / Path(str(filename)).name, *args, **kwargs)

    RealDocument.save = save

    original_init = RealDocument.__init__

    def init(self, filename=None, *args, **kwargs):
        original_init(self, None if filename is None else DATA_FILE, *args, **kwargs)

    RealDocument.__init__ = init
    return RealDocument


def open_from_data(file, *args, **kwargs):
    """open() replacement that resolves relative filenames in tests/data."""
    if isinstance(file, (str, os.PathLike)) and not Path(file).is_absolute():
        file = DATA_DIR / file
    return open(file, *args, **kwargs)


def stored_and_loaded(code: str) -> tuple[set[str], set[str]]:
    tree = ast.parse(code)
    stored, loaded = set(), set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            (stored if isinstance(node.ctx, ast.Store) else loaded).add(node.id)
    return stored, loaded


def source_of(example: Example) -> str:
    if not example.is_session:
        return example.text
    parser = doctest.DocTestParser()
    return "\n".join(e.source for e in parser.get_examples(example.text))


def build_namespace(example: Example, document_cls) -> dict:
    ns = {name: getattr(numbers_parser, name) for name in dir(numbers_parser)}
    ns.update({"Document": document_cls, "open": open_from_data, "datetime": datetime})
    try:
        src = source_of(example)
        stored, loaded = stored_and_loaded(src)
    except SyntaxError:
        print(f"FAIL (syntax error): {src}")
        return ns
    needed = {"doc", "sheets", "sheet", "tables", "table"} & loaded - stored
    if not needed:
        return ns
    doc = document_cls("mydoc.numbers")
    ns["doc"] = doc
    ns["sheets"] = doc.sheets
    ns["sheet"] = doc.sheets[0]
    ns["tables"] = doc.sheets[0].tables
    ns["table"] = doc.sheets[0].tables["Examples"]
    return ns


class AddressChecker(doctest.OutputChecker):
    def check_output(self, want, got, optionflags):
        return super().check_output(
            ADDRESS_RE.sub("0xADDR", want), ADDRESS_RE.sub("0xADDR", got), optionflags
        )


def run_session(example: Example, ns: dict) -> str | None:
    parser = doctest.DocTestParser()
    test = parser.get_doctest(example.text, ns, example.where, str(example.path), example.line)
    out = io.StringIO()
    runner = doctest.DocTestRunner(
        checker=AddressChecker(), verbose=False, optionflags=doctest.NORMALIZE_WHITESPACE
    )
    runner.run(test, out=out.write, clear_globs=False)
    return out.getvalue() if runner.failures else None


def run_code(example: Example, ns: dict) -> str | None:
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            exec(compile(example.text, example.where, "exec"), ns)  # noqa: S102
    except BaseException:
        return traceback.format_exc(limit=-2)
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    if not DATA_FILE.exists():
        print(f"{DATA_FILE} missing")
        return 1

    examples = find_examples()
    failures = 0
    with tempfile.TemporaryDirectory() as tmp:
        document_cls = make_document_shim(Path(tmp))
        for ex in examples:
            kind = "session" if ex.is_session else "code"
            ns = build_namespace(ex, document_cls)
            err = run_session(ex, ns) if ex.is_session else run_code(ex, ns)
            if err is not None:
                failures += 1
            if args.verbose:
                if err is not None:
                    print(f"FAIL {ex.where} ({kind})")
                    print(textwrap.indent(err.rstrip(), "     "))
                else:
                    print(f"PASS {ex.where} ({kind})")
                    print(textwrap.indent(ex.text, "     "))

    argv0 = os.path.basename(__file__)
    print(f"{argv0}: {len(examples)} examples checked, {failures} problems")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
