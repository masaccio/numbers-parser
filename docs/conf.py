import os
import sys

from numbers_parser import _get_version

sys.path.insert(0, os.path.abspath("../"))

templates_path = ["_templates"]
exclude_patterns = ["build", "Thumbs.db", ".DS_Store"]
language = "en"
master_doc = "index"
project = "numbers-parser"
copyright = "Copyright Jon Connell under MIT license"

# sphinx.ext.napoleon configuration
napoleon_google_docstring = False
napoleon_numpy_docstring = True
napoleon_include_special_with_doc = False
napoleon_use_admonition_for_examples = True
napoleon_use_admonition_for_notes = True

# sphinx_nefertiti options
html_theme = "sphinx_nefertiti"
html_style = ["custom.css", "nftt-pygments.min.css"]
pygments_style = "pastie"
pygments_dark_style = "dracula"
html_theme_options = {
    # "documentation_font": "Open Sans",
    # "monospace_font": "Ubuntu Mono",
    # "monospace_font_size": "1.1rem",
    # "logo": "docs/logo.svg",
    # "logo_alt": "numbers-parser",
    "repository_url": "https://github.com/masaccio/numbers-parser",
    "repository_name": "masaccio/numbers-parser",
    "current_version": _get_version(),
    "footer_links": ",".join(
        [
            "Documentation|https://masaccio.github.io/numbers-parser/",
            "Package|https://pypi.org/project/numbers-parser/",
            "Repository|https://github.com/masaccio/numbers-parser",
            "Issues|https://github.com/masaccio/numbers-parser/issues",
        ]
    ),
    # "show_colorset_choices": True,
}

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.napoleon",
    "enum_tools.autoenum",
]


def setup_extensions(app, docname, source):
    if app.builder.name == "html":
        extensions.append("sphinx_nefertiti")


def setup(app):
    app.connect("source-read", setup_extensions)
