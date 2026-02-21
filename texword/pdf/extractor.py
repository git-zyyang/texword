"""PDF 结构提取器 — 从 PDF 中提取文本、公式、表格、图片。"""

import os
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


class BlockType(Enum):
    TEXT = "text"
    HEADING = "heading"
    EQUATION = "equation"
    TABLE = "table"
    FIGURE = "figure"
    REFERENCE = "reference"
    ABSTRACT = "abstract"
    TITLE = "title"
    AUTHOR = "author"


@dataclass
class ContentBlock:
    """PDF 中提取的内容块。"""
    type: BlockType
    text: str = ""
    latex: str = ""  # OCR 识别的 LaTeX（公式用）
    image_path: str = ""  # 图片路径
    page: int = 0
    bbox: tuple = ()  # (x0, y0, x1, y1)
    confidence: float = 1.0
    metadata: dict = field(default_factory=dict)


class PDFExtractor:
    """
    PDF 结构提取器。

    从学术论文 PDF 中提取结构化内容块，
    识别标题、摘要、正文、公式、表格、图片、参考文献。
    """

    def __init__(self, pdf_path: str):
        if not HAS_FITZ:
            raise ImportError("PyMuPDF is required: pip install PyMuPDF")
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.blocks: list[ContentBlock] = []

    def extract_text_blocks(self) -> list[ContentBlock]:
        """提取所有文本块（基础版，不含 OCR）。"""
        blocks = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text_dict = page.get_text("dict")

            for block in text_dict.get("blocks", []):
                if block["type"] == 0:  # text block
                    text = ""
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text += span["text"]
                        text += "\n"
                    text = text.strip()
                    if not text:
                        continue

                    bbox = tuple(block["bbox"])
                    block_type = self._classify_block(
                        text, bbox, page_num,
                        spans=block.get("lines", [{}])[0].get("spans", [])
                    )
                    blocks.append(ContentBlock(
                        type=block_type,
                        text=text,
                        page=page_num,
                        bbox=bbox,
                    ))

        self.blocks = blocks
        return blocks

    def extract_figures(self, output_dir: str) -> list[ContentBlock]:
        """提取 PDF 中的图片。"""
        os.makedirs(output_dir, exist_ok=True)
        figures = []

        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            images = page.get_images(full=True)

            for img_idx, img_info in enumerate(images):
                xref = img_info[0]
                try:
                    pix = fitz.Pixmap(self.doc, xref)
                    if pix.n > 4:  # CMYK → RGB
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    img_name = f"page{page_num + 1}_img{img_idx + 1}.png"
                    img_path = os.path.join(output_dir, img_name)
                    pix.save(img_path)

                    figures.append(ContentBlock(
                        type=BlockType.FIGURE,
                        image_path=img_path,
                        page=page_num,
                        metadata={"xref": xref},
                    ))
                except Exception as e:
                    print(f"  [WARN] Failed to extract image p{page_num + 1}: {e}")

        return figures

    def _classify_block(self, text: str, bbox: tuple,
                        page_num: int, spans: list = None) -> BlockType:
        """启发式分类文本块类型。"""
        # 首页大字体 → 标题
        if page_num == 0 and spans:
            font_size = max((s.get("size", 12) for s in spans), default=12)
            if font_size >= 16:
                return BlockType.TITLE
            if font_size >= 13 and bbox[1] < 300:
                return BlockType.AUTHOR

        text_lower = text.lower().strip()

        # Abstract
        if text_lower.startswith("abstract"):
            return BlockType.ABSTRACT

        # Headings (numbered sections)
        import re
        if re.match(r"^\d+\.?\s+[A-Z]", text) and len(text) < 100:
            return BlockType.HEADING

        # References
        if re.match(r"^[A-Z][a-z]+,?\s.*\(\d{4}\)", text):
            return BlockType.REFERENCE

        return BlockType.TEXT

    def close(self):
        self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
