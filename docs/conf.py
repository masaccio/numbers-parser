import os
import sys
from typing import Optional

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.napoleon",
    "sphinx.ext.autosectionlabel",
    "sphinx_autodoc_typehints",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Jon Connell"


def format_optional(annotation, config):
    if annotation == Optional[str]:
        return "``str`` *optional*"
    elif annotation == Optional[int]:
        return "``int`` *optional*"
    elif annotation == Optional[float]:
        return "``float`` *optional*"
    elif annotation == Optional[bool]:
        return "``bool`` *optional*"
    if "Optional" in str(annotation):
        print(f"format_optional: unknown optional annotation '{annotation}'")
    return None


# sphinx_autodoc_typehints options
autodoc_typehints = "both"
typehints_use_signature = True
typehints_use_signature_return = True
simplify_optional_unions = False
always_document_param_types = True
typehints_defaults = "comma"
typehints_formatter = format_optional

autodoc_default_options = {
    "member-order": "bysource",
}

sys.path.insert(0, os.path.abspath("../"))
