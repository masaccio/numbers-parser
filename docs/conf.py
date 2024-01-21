import os
import sys

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.todo",
    "sphinx_toolbox.more_autodoc.typehints",
    "sphinx_autodoc_typehints",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Jon Connell"

autodoc_typehints = "both"

autodoc_default_options = {
    "member-order": "bysource",
    "members": True,
    "show-inheritance": False,
    "hide_none_rtype": True,
}

sys.path.insert(0, os.path.abspath("../"))
