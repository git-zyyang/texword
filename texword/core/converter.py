"""主转换流程编排。"""

import os
import re
import shutil
from pathlib import Path

from texword.core.style import StyleConfig
from texword.latex.preprocessor import preprocess
from texword.latex.pandoc_bridge import pandoc_convert
from texword.latex.postprocessor import postprocess


def extract_title_short(tex_path: str) -> str:
    """从 LaTeX 文件提取短标题（用于页眉）。"""
    with open(tex_path, "r", encoding="utf-8") as f:
        tex = f.read()
    m = re.search(r"\\title\{(?:\\textbf\{)?([^}]+)", tex)
    if m:
        title = m.group(1).strip()
        title = title.split("\\\\")[0].strip()
        if len(title) > 60:
            title = title[:57] + "..."
        return title
    return ""


def convert(tex_path: str, output_path: str = None,
            cfg: StyleConfig = None, cleanup: bool = True):
    """
    主转换函数：LaTeX → Word

    Args:
        tex_path: 输入 .tex 文件路径
        output_path: 输出 .docx 文件路径（默认同名）
        cfg: 样式配置
        cleanup: 是否清理临时文件
    """
    if cfg is None:
        cfg = StyleConfig()

    if output_path is None:
        output_path = str(Path(tex_path).with_suffix(".docx"))

    print(f"{'=' * 60}")
    print(f"TexWord — LaTeX → Word 高质量转换器")
    print(f"{'=' * 60}")
    print(f"输入: {tex_path}")
    print(f"输出: {output_path}")

    title_short = extract_title_short(tex_path)
    if title_short:
        print(f"标题: {title_short}")

    processed_tex, tmp_dir = preprocess(tex_path)

    try:
        raw_docx = pandoc_convert(processed_tex, tmp_dir)
        postprocess(raw_docx, output_path, cfg, title_short)

        print(f"\n{'=' * 60}")
        print(f"转换完成!")
        print(f"{'=' * 60}")
    finally:
        if cleanup:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        else:
            print(f"\n[DEBUG] 临时文件保留在: {tmp_dir}")
