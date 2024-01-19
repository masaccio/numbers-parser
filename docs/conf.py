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
    "show-inheritance": True,
}

sys.path.insert(0, os.path.abspath("../"))

# extensions.append("sphinx_autodoc_typehints")
