"""通用 API 引擎 — 支持 Claude Vision、GPT-4V 等。"""

import os
import base64
from texword.pdf.ocr_engine import OCREngine, OCRResult


class APIGenericEngine(OCREngine):
    """
    通用 Vision API 引擎。

    支持任何兼容 OpenAI Chat Completions 格式的 API：
    - Claude Vision (via Anthropic API)
    - GPT-4V (via OpenAI API)
    - 其他兼容 API

    Usage:
        engine = APIGenericEngine(
            api_key="your_key",
            base_url="https://api.openai.com/v1/chat/completions",
            model="gpt-4o",
        )
    """

    def __init__(self, api_key: str = None, base_url: str = None,
                 model: str = "gpt-4o"):
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        self.base_url = base_url or "https://api.openai.com/v1/chat/completions"
        self.model = model

    @property
    def name(self) -> str:
        return f"API-Generic ({self.model})"

    def recognize_text(self, image_path: str) -> list[OCRResult]:
        response = self._call_api(image_path, "Extract all text from this image, preserving formatting.")
        return [OCRResult(text=response, confidence=0.85)]

    def recognize_equation(self, image_path: str) -> OCRResult:
        response = self._call_api(
            image_path,
            "Recognize the math equation in this image. Output only LaTeX code, nothing else."
        )
        return OCRResult(latex=response, confidence=0.85, block_type="equation")

    def recognize_table(self, image_path: str) -> OCRResult:
        response = self._call_api(
            image_path,
            "Recognize the table in this image. Output LaTeX tabular format preserving all data."
        )
        return OCRResult(latex=response, confidence=0.8, block_type="table")

    def _call_api(self, image_path: str, prompt: str) -> str:
        import json
        import urllib.request

        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        payload = {
            "model": self.model,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
                    {"type": "text", "text": prompt},
                ]
            }],
            "max_tokens": 4096,
        }

        req = urllib.request.Request(
            self.base_url,
            data=json.dumps(payload).encode(),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}",
            },
        )

        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())

        return result["choices"][0]["message"]["content"]
