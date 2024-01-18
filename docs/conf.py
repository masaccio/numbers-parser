import os
import sys

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.todo",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Jon Connell"


sys.path.insert(0, os.path.abspath("../"))


def skip(app, what, name, obj, would_skip, options):  # noqa: PLR0913
    if name in ("__init__",):
        return False
    return would_skip


def setup(app):
    app.connect("autodoc-skip-member", skip)


extensions.append("sphinx_autodoc_typehints")
