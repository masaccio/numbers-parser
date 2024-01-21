import os
import sys

# from sphinx_toolbox.more_autodoc.typehints import hide_non_rtype

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

# typehints_use_signature = True
# typehints_use_signature_return = True
# hide_none_rtype = True
# typehints_document_rtype = False

autodoc_default_options = {
    # "all-typevars": True,
    # "autoclass-content": "class",
    "member-order": "bysource",
    "members": True,
    "show-inheritance": False,
}

# class-doc-from
# exclude-members
# ignore-module-all
# imported-members
# inherited-members
# member-order
# members
# no-value
# private-members
# show-inheritance
# special-members
# undoc-members

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


# def setup(app):
#     app.connect("autodoc-skip-member", autodoc_skip_member)
