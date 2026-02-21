"""TexWord 基本测试。"""

import os
import tempfile
import pytest

from texword.core.style import StyleConfig
from texword.latex.preprocessor import (
    resolve_inputs, strip_comments, fix_latex_for_pandoc
)


class TestStyleConfig:
    def test_defaults(self):
        cfg = StyleConfig()
        assert cfg.font_body == "Times New Roman"
        assert cfg.font_size_body == 12
        assert cfg.line_spacing == 2.0
        assert cfg.margin_top == 2.54

    def test_override(self):
        cfg = StyleConfig()
        cfg.font_size_body = 11
        cfg.line_spacing = 1.5
        assert cfg.font_size_body == 11
        assert cfg.line_spacing == 1.5


class TestPreprocessor:
    def test_strip_comments(self):
        tex = r"Hello % this is a comment" + "\n" + r"World \% not a comment"
        result = strip_comments(tex)
        assert "this is a comment" not in result
        assert r"\% not a comment" in result

    def test_resolve_inputs_missing_file(self, capsys):
        tex = r"\input{nonexistent_file}"
        result = resolve_inputs(tex, "/tmp")
        assert result == r"\input{nonexistent_file}"
        captured = capsys.readouterr()
        assert "WARN" in captured.out

    def test_resolve_inputs_with_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            # Create a sub-file
            sub_path = os.path.join(tmp, "sub.tex")
            with open(sub_path, "w") as f:
                f.write("SUB CONTENT")
            tex = r"\input{sub}"
            result = resolve_inputs(tex, tmp)
            assert "SUB CONTENT" in result

    def test_fix_latex_threeparttable(self):
        tex = r"\begin{threeparttable}" + "\n" + r"\end{threeparttable}"
        result = fix_latex_for_pandoc(tex)
        assert "threeparttable" not in result

    def test_fix_latex_amsmath_injection(self):
        tex = r"\begin{document}" + "\nHello"
        result = fix_latex_for_pandoc(tex)
        assert r"\usepackage{amsmath}" in result

    def test_fix_latex_amsmath_no_duplicate(self):
        tex = r"\usepackage{amsmath}" + "\n" + r"\begin{document}"
        result = fix_latex_for_pandoc(tex)
        assert result.count("amsmath") == 1


class TestCLI:
    def test_version(self):
        from texword import __version__
        assert __version__ == "0.1.0"

    def test_pdf_extractor_file_not_found(self):
        from texword.pdf.extractor import PDFExtractor
        with pytest.raises(FileNotFoundError):
            PDFExtractor("nonexistent.pdf")

    def test_pdf_block_types(self):
        from texword.pdf.extractor import BlockType
        assert BlockType.EQUATION.value == "equation"
        assert BlockType.TABLE.value == "table"

    def test_ocr_engine_interface(self):
        from texword.pdf.ocr_engine import OCREngine
        # OCREngine is abstract, can't instantiate
        with pytest.raises(TypeError):
            OCREngine()
