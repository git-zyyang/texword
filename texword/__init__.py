"""TexWord — High-quality academic paper converter: LaTeX/PDF → editable Word."""

__version__ = "0.1.0"

from texword.core.converter import convert
from texword.core.style import StyleConfig

__all__ = ["convert", "StyleConfig", "__version__"]
