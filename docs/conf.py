import os
import sys

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.todo",
    "sphinx_autodoc_typehints",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Jon Connell"

autodoc_typehints = "signature"

autodoc_default_options = {
    "hide_none_rtype": True,
    "all_typevars": True,
    "autoclass_content": "class",
    "member-order": "bysource",
    "members": True,
    "show-inheritance": False,
}

sys.path.insert(0, os.path.abspath("../"))


def autodoc_skip_member(app, what, name, obj, skip, options):  # noqa: PLR0913
    if skip:
        return True
    if what == "class":
        if name in ["Table", "Sheet"]:
            return True
    elif what == "method":
        if name in ["Table.__init__", "Sheet.__init__"]:
            return True
    return False


def setup(app):
    app.connect("autodoc-skip-member", autodoc_skip_member)
