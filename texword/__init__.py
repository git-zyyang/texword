"""TexWord — Academic document converter: LaTeX/Markdown/PDF → Word/PDF."""

__version__ = "0.2.0"

from texword.core.converter import convert
from texword.core.style import StyleConfig
from texword.markdown.converter import md_to_pdf, md_to_docx

__all__ = ["convert", "md_to_pdf", "md_to_docx", "StyleConfig", "__version__"]
