"""TexWord CLI 入口。"""

import argparse
import sys

from texword import __version__
from texword.core.style import StyleConfig


def _infer_output(input_path: str, output_path: str) -> str:
    """根据输入格式和 -o 参数推断输出路径。"""
    from pathlib import Path
    if output_path:
        return output_path
    ext = Path(input_path).suffix.lower()
    # Markdown 默认输出 PDF，其余默认输出 Word
    default_suffix = ".pdf" if ext == ".md" else ".docx"
    return str(Path(input_path).with_suffix(default_suffix))


def main():
    parser = argparse.ArgumentParser(
        prog="texword",
        description="TexWord — 学术文档格式转换器 (LaTeX/Markdown/PDF → Word/PDF)",
    )
    parser.add_argument("input",
                        help="输入文件路径 (.tex / .md / .pdf)")
    parser.add_argument("-o", "--output",
                        help="输出文件路径 (.docx 或 .pdf)")
    parser.add_argument("--font-size", type=int, default=12,
                        help="正文字号 (默认 12)")
    parser.add_argument("--font", default="Times New Roman",
                        help="正文字体 (默认 Times New Roman)")
    parser.add_argument("--line-spacing", type=float, default=2.0,
                        help="行距倍数 (默认 2.0)")
    parser.add_argument("--css",
                        help="自定义 CSS 样式 (Markdown→PDF 时使用)")
    parser.add_argument("--no-cleanup", action="store_true",
                        help="保留临时文件（调试用）")
    parser.add_argument("--ocr-key",
                        help="OCR API key (DeepSeek-OCR-2)")
    parser.add_argument("--ocr-url", default="https://aiping.cn/api/v1",
                        help="OCR API base URL")
    parser.add_argument("--version", action="version",
                        version=f"texword {__version__}")

    args = parser.parse_args()

    input_ext = args.input.lower().rsplit(".", 1)[-1] if "." in args.input else ""
    output = _infer_output(args.input, args.output)
    output_ext = output.lower().rsplit(".", 1)[-1] if "." in output else ""

    # ── Markdown 路线 ──
    if input_ext == "md":
        from texword.markdown.converter import md_to_pdf, md_to_docx

        if output_ext == "pdf":
            md_to_pdf(args.input, output, css_path=args.css)
        else:
            cfg = StyleConfig()
            cfg.font_body = args.font
            cfg.font_size_body = args.font_size
            cfg.line_spacing = args.line_spacing
            md_to_docx(args.input, output, cfg)
        return

    # ── PDF 路线 ──
    if input_ext == "pdf":
        from texword.pdf.extractor import PDFExtractor
        from texword.pdf.assembler import PDFAssembler
        from pathlib import Path
        import tempfile
        import os

        print(f"{'=' * 60}")
        print(f"TexWord — PDF → Word 转换器")
        print(f"{'=' * 60}")
        print(f"输入: {args.input}")
        print(f"输出: {output}")

        ocr_engine = None
        ocr_key = args.ocr_key or os.environ.get("DEEPSEEK_OCR_API_KEY")
        if ocr_key:
            try:
                from texword.pdf.engines.deepseek_ocr import DeepSeekOCREngine
                ocr_engine = DeepSeekOCREngine(
                    api_key=ocr_key, base_url=args.ocr_url)
                print(f"OCR: {ocr_engine.name}")
            except Exception as e:
                print(f"  [WARN] OCR 初始化失败: {e}")
        else:
            print("OCR: 未配置 (仅提取文本，公式将丢失)")
            print("  设置 DEEPSEEK_OCR_API_KEY 或 --ocr-key 启用公式识别")

        with PDFExtractor(args.input) as ext:
            blocks = ext.extract_text_blocks()
            tmp_dir = tempfile.mkdtemp(prefix="texword_pdf_")
            figures = ext.extract_figures(tmp_dir)
            blocks.extend(figures)

        print(f"  提取: {len(blocks)} 个内容块, {len(figures)} 张图片")

        cfg = StyleConfig()
        cfg.font_body = args.font
        cfg.font_size_body = args.font_size
        cfg.line_spacing = args.line_spacing

        assembler = PDFAssembler(cfg)
        assembler.assemble(blocks, output)
        print(f"\n{'=' * 60}")
        print(f"转换完成!")
        print(f"{'=' * 60}")
        return

    # ── LaTeX 路线（默认） ──
    from texword.core.converter import convert

    cfg = StyleConfig()
    cfg.font_body = args.font
    cfg.font_size_body = args.font_size
    cfg.line_spacing = args.line_spacing

    convert(args.input, output, cfg, cleanup=not args.no_cleanup)


if __name__ == "__main__":
    main()
