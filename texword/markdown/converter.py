"""Markdown → PDF / Word 转换器。"""

import os
import platform
import subprocess
import tempfile
from pathlib import Path

from texword.core.style import StyleConfig
from texword.latex.postprocessor import postprocess

# 内置 CSS 路径
_BUILTIN_CSS = Path(__file__).parent / "styles" / "academic.css"


def _pandoc_env() -> dict:
    """构建 pandoc 子进程环境变量（macOS 需要 gobject 库路径）。"""
    env = os.environ.copy()
    if platform.system() == "Darwin":
        brew_lib = "/opt/homebrew/lib"
        if os.path.isdir(brew_lib):
            env["DYLD_LIBRARY_PATH"] = brew_lib + ":" + env.get(
                "DYLD_LIBRARY_PATH", "")
    return env


def md_to_pdf(input_path: str, output_path: str = None,
              css_path: str = None) -> str:
    """Markdown → PDF（pandoc + weasyprint）。

    Args:
        input_path: 输入 .md 文件
        output_path: 输出 .pdf 文件（默认同名）
        css_path: 自定义 CSS（默认使用内置学术样式）

    Returns:
        输出文件路径
    """
    if output_path is None:
        output_path = str(Path(input_path).with_suffix(".pdf"))

    css = css_path or str(_BUILTIN_CSS)

    print(f"\n[Markdown → PDF]")
    print(f"  输入: {input_path}")
    print(f"  样式: {Path(css).name}")

    cmd = [
        "pandoc", input_path,
        "-o", output_path,
        "--pdf-engine=weasyprint",
        f"--css={css}",
    ]

    result = subprocess.run(
        cmd, capture_output=True, text=True,
        timeout=120, env=_pandoc_env())

    if result.returncode != 0:
        raise RuntimeError(
            f"pandoc 转换失败:\n{result.stderr[:500]}")

    # 忽略 weasyprint 的 CSS 警告
    if result.stderr:
        warnings = [l for l in result.stderr.splitlines()
                    if not l.startswith("WARNING:")]
        if warnings:
            print(f"  [WARN] {chr(10).join(warnings[:3])}")

    size_kb = os.path.getsize(output_path) / 1024
    print(f"  完成 → {output_path} ({size_kb:.0f} KB)")
    return output_path


def md_to_docx(input_path: str, output_path: str = None,
               cfg: StyleConfig = None) -> str:
    """Markdown → Word（pandoc + postprocess 格式化）。

    Args:
        input_path: 输入 .md 文件
        output_path: 输出 .docx 文件（默认同名）
        cfg: 样式配置

    Returns:
        输出文件路径
    """
    if cfg is None:
        cfg = StyleConfig()
    if output_path is None:
        output_path = str(Path(input_path).with_suffix(".docx"))

    print(f"\n[Markdown → Word]")
    print(f"  输入: {input_path}")

    # Phase 1: pandoc 转为原始 docx
    tmp_dir = tempfile.mkdtemp(prefix="texword_md_")
    raw_docx = os.path.join(tmp_dir, "raw_output.docx")

    cmd = [
        "pandoc", input_path,
        "-o", raw_docx,
        "-f", "markdown",
        "-t", "docx",
        "--wrap=none",
    ]

    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0 or not os.path.exists(raw_docx):
        raise RuntimeError(
            f"pandoc 转换失败:\n{result.stderr[:500]}")

    # Phase 2: postprocess 格式化
    postprocess(raw_docx, output_path, cfg, title_short="")

    # 清理
    import shutil
    shutil.rmtree(tmp_dir, ignore_errors=True)

    size_kb = os.path.getsize(output_path) / 1024
    print(f"  完成 → {output_path} ({size_kb:.0f} KB)")
    return output_path
