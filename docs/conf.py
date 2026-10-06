import os
import sys

import docutils.nodes

from numbers_parser import _get_version

sys.path.insert(0, os.path.abspath("../"))  # noqa: PTH100

GITHUB = "https://github.com/masaccio/numbers-parser"
PAGES = "https://masaccio.github.io/numbers-parser"

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.napoleon",
    "sphinx.ext.extlinks",
    "enum_tools.autoenum",
    "sphinx_copybutton",
]
# Standard Sphinx configuration
templates_path = ["_templates"]
exclude_patterns = ["build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Copyright Jon Connell under MIT license"  # noqa: A001

# sphinx.ext.napoleon configuration
napoleon_google_docstring = False
napoleon_numpy_docstring = True
napoleon_include_special_with_doc = False
# napoleon_use_admonition_for_examples = True
napoleon_use_admonition_for_notes = True

# sphinx_nefertiti options
html_theme = "sphinx_nefertiti"
html_style = ["custom.css", "nftt-pygments.min.css"]
pygments_style = "pastie"
pygments_dark_style = "dracula"
html_theme_options = {
    "monospace_font": "Ubuntu Sans",
    "monospace_font_size": "1.1rem",
    "repository_url": GITHUB,
    "repository_name": "masaccio/numbers-parser",
    "current_version": _get_version(),
    "footer_links": [
        {
            "text": "Documentation",
            "link": f"{PAGES}",
        },
        {
            "text": "Repository",
            "link": f"{GITHUB}",
        },
        {
            "text": "Issues",
            "link": f"{GITHUB}/issues",
        },
    ],
    "show_colorset_choices": True,
}


# sphinx_copybutton options
copybutton_prompt_text = ">>> "
copybutton_line_continuation_character = "\\"

extlinks = {
    "github": (f"{GITHUB}/%s", None),
    "pages": (f"{PAGES}/%s", None),
}


def create_github_role(base_path):
    """Factory to create roles that link to specific GitHub repository paths."""

    def role(name, rawtext, text, *args):
        url = f"{GITHUB}/blob/main{base_path}{text}"
        node = docutils.nodes.reference("", "", internal=False, refuri=url)
        node += docutils.nodes.literal("", text)

        return [node], []

    return role


def setup_extensions(app, *args):
    if app.builder.name == "html":
        extensions.append("sphinx_nefertiti")


def setup(app):
    app.add_role("src_root", create_github_role("/"))
    app.add_role("src_pkg", create_github_role("/src/numbers_parser/"))
    app.add_role("src_gen", create_github_role("/src/numbers_parser/generated/"))
    app.add_role("src_proto", create_github_role("/src/protos/"))
    app.connect("source-read", setup_extensions)
