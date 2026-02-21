"""OCR 引擎抽象层 — 统一接口调用不同 OCR 模型。"""

from abc import ABC, abstractmethod
from dataclasses import dataclass


@dataclass
class OCRResult:
    """OCR 识别结果。"""
    text: str = ""
    latex: str = ""  # 公式的 LaTeX 表示
    confidence: float = 0.0
    block_type: str = "text"  # text / equation / table


class OCREngine(ABC):
    """OCR 引擎基类。所有引擎实现此接口。"""

    @abstractmethod
    def recognize_text(self, image_path: str) -> list[OCRResult]:
        """识别图片中的文本。"""
        ...

    @abstractmethod
    def recognize_equation(self, image_path: str) -> OCRResult:
        """识别图片中的数学公式，返回 LaTeX。"""
        ...

    @abstractmethod
    def recognize_table(self, image_path: str) -> OCRResult:
        """识别图片中的表格，返回结构化数据。"""
        ...

    @property
    @abstractmethod
    def name(self) -> str:
        """引擎名称。"""
        ...
