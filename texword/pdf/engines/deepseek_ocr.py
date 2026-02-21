"""DeepSeek-OCR-2 引擎 — 学术公式识别，兼容 OpenAI API 格式。

需要: pip install openai
设置: export DEEPSEEK_OCR_API_KEY=your_key
"""

import os
import base64
from texword.pdf.ocr_engine import OCREngine, OCRResult


class DeepSeekOCREngine(OCREngine):
    """
    DeepSeek-OCR-2 — 学术论文公式/表格识别引擎。

    通过 OpenAI 兼容 API 调用，支持流式输出。

    Usage:
        engine = DeepSeekOCREngine(api_key="your_key")
        result = engine.recognize_equation("equation.png")
        print(result.latex)
    """

    def __init__(self, api_key: str = None,
                 base_url: str = "https://aiping.cn/api/v1",
                 model: str = "DeepSeek-OCR-2"):
        self.api_key = api_key or os.environ.get("DEEPSEEK_OCR_API_KEY")
        if not self.api_key:
            raise ValueError(
                "DeepSeek OCR API key required. "
                "Set DEEPSEEK_OCR_API_KEY env var or pass api_key."
            )
        self.base_url = base_url
        self.model = model
        self._client = None

    @property
    def client(self):
        if self._client is None:
            from openai import OpenAI
            self._client = OpenAI(
                base_url=self.base_url,
                api_key=self.api_key,
            )
        return self._client

    @property
    def name(self) -> str:
        return "DeepSeek-OCR-2"

    def recognize_text(self, image_path: str) -> list[OCRResult]:
        response = self._call(image_path, "识别图片中的所有文本内容，保持原始格式和排版。")
        return [OCRResult(text=response, confidence=0.9)]

    def recognize_equation(self, image_path: str) -> OCRResult:
        response = self._call(
            image_path,
            "识别图片中的数学公式，输出标准 LaTeX 格式。只输出 LaTeX 代码，不要其他解释文字。"
        )
        return OCRResult(latex=response.strip(), confidence=0.9, block_type="equation")

    def recognize_table(self, image_path: str) -> OCRResult:
        response = self._call(
            image_path,
            "识别图片中的表格，输出 LaTeX tabular 格式。保持所有数据、对齐和格式。"
        )
        return OCRResult(latex=response.strip(), confidence=0.85, block_type="table")

    def recognize_page(self, image_path: str) -> str:
        """识别整页 PDF 内容，返回完整文本+公式的混合输出。"""
        return self._call(
            image_path,
            "识别这一页学术论文的完整内容。"
            "文本部分直接输出，数学公式用 LaTeX 格式（行内公式用 $...$，"
            "独立公式用 $$...$$）。保持原始段落结构。"
        )

    def _call(self, image_path: str, prompt: str) -> str:
        """调用 DeepSeek-OCR API。"""
        # 读取图片并编码
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        # 判断图片类型
        ext = os.path.splitext(image_path)[1].lower()
        mime = {"png": "image/png", "jpg": "image/jpeg",
                "jpeg": "image/jpeg"}.get(ext.lstrip("."), "image/png")

        response = self.client.chat.completions.create(
            model=self.model,
            stream=False,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime};base64,{img_b64}"
                        },
                    },
                    {
                        "type": "text",
                        "text": prompt,
                    },
                ],
            }],
        )

        return response.choices[0].message.content

    def _call_stream(self, image_path: str, prompt: str):
        """流式调用，逐块返回内容。"""
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        ext = os.path.splitext(image_path)[1].lower()
        mime = {"png": "image/png", "jpg": "image/jpeg",
                "jpeg": "image/jpeg"}.get(ext.lstrip("."), "image/png")

        response = self.client.chat.completions.create(
            model=self.model,
            stream=True,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime};base64,{img_b64}"},
                    },
                    {"type": "text", "text": prompt},
                ],
            }],
        )

        for chunk in response:
            if not getattr(chunk, "choices", None):
                continue
            content = getattr(chunk.choices[0].delta, "content", None)
            if content:
                yield content
