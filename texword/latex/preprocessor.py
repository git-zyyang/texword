"""LaTeX 预处理：展开 \\input、转换 PDF 图片、修复 pandoc 兼容性。"""

import os
import re
import shutil
import tempfile
from pathlib import Path

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False


def resolve_inputs(tex: str, base_dir: str) -> str:
    """递归展开所有 \\input{} 命令。"""
    def replacer(m):
        fname = m.group(1)
        if not fname.endswith(".tex"):
            fname += ".tex"
        fpath = os.path.join(base_dir, fname)
        if os.path.exists(fpath):
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()
            sub_dir = os.path.dirname(fpath)
            return resolve_inputs(content, sub_dir or base_dir)
        print(f"  [WARN] \\input file not found: {fpath}")
        return m.group(0)
    return re.sub(r"\\input\{([^}]+)\}", replacer, tex)


def convert_pdf_figures(tex: str, base_dir: str, tmp_dir: str) -> str:
    """将 \\includegraphics 引用的 PDF 图片转为 PNG（300 DPI）。"""
    graphics_path = base_dir
    gp_match = re.search(r"\\graphicspath\{\{([^}]+)\}\}", tex)
    if gp_match:
        graphics_path = os.path.join(base_dir, gp_match.group(1))

    def replacer(m):
        full_match = m.group(0)
        opts = m.group(1) or ""
        fname = m.group(2)

        # 查找图片文件
        for ext_try in ["", ".pdf", ".png", ".jpg", ".jpeg"]:
            candidate = os.path.join(graphics_path, fname + ext_try)
            if os.path.exists(candidate):
                fpath = candidate
                break
        else:
            for ext_try in ["", ".pdf", ".png", ".jpg", ".jpeg"]:
                candidate = os.path.join(base_dir, fname + ext_try)
                if os.path.exists(candidate):
                    fpath = candidate
                    break
            else:
                print(f"  [WARN] Figure not found: {fname}")
                return full_match

        if fpath.lower().endswith(".pdf"):
            if not HAS_FITZ:
                print(f"  [WARN] PyMuPDF not available, cannot convert {fpath}")
                return full_match
            png_name = Path(fpath).stem + ".png"
            png_path = os.path.join(tmp_dir, png_name)
            try:
                doc = fitz.open(fpath)
                page = doc[0]
                mat = fitz.Matrix(300 / 72, 300 / 72)
                pix = page.get_pixmap(matrix=mat)
                pix.save(png_path)
                doc.close()
                print(f"  [OK] PDF→PNG: {Path(fpath).name} → {png_name}")
                return f"\\includegraphics[{opts}]{{{png_path}}}"
            except Exception as e:
                print(f"  [WARN] PDF conversion failed for {fpath}: {e}")
                return full_match
        else:
            dst = os.path.join(tmp_dir, os.path.basename(fpath))
            if not os.path.exists(dst):
                shutil.copy2(fpath, dst)
            return f"\\includegraphics[{opts}]{{{dst}}}"

    return re.sub(
        r"\\includegraphics\[([^\]]*)\]\{([^}]+)\}",
        replacer, tex
    )


def strip_comments(tex: str) -> str:
    """移除 LaTeX 注释（保留 \\%）。"""
    lines = []
    for line in tex.split("\n"):
        cleaned = re.sub(r"(?<!\\)%.*$", "", line)
        lines.append(cleaned)
    return "\n".join(lines)


def fix_latex_for_pandoc(tex: str) -> str:
    """修复 pandoc 不能正确处理的 LaTeX 结构。"""
    # threeparttable
    tex = re.sub(r"\\begin\{threeparttable\}", "", tex)
    tex = re.sub(r"\\end\{threeparttable\}", "", tex)
    tex = re.sub(r"\\begin\{tablenotes\}.*?\\end\{tablenotes\}",
                 "", tex, flags=re.DOTALL)
    # fancyhdr
    tex = re.sub(r"\\usepackage\{fancyhdr\}.*?\n", "\n", tex)
    tex = re.sub(r"\\pagestyle\{fancy\}.*?\n", "\n", tex)
    tex = re.sub(r"\\fancyhf\{\}.*?\n", "\n", tex)
    tex = re.sub(r"\\renewcommand\{\\headrulewidth\}.*?\n", "\n", tex)
    tex = re.sub(r"\\[rl]head\{.*?\}\n", "\n", tex)
    # titlesec
    tex = re.sub(r"\\titleformat\{.*?\}.*?\n", "\n", tex)
    tex = re.sub(r"\\usepackage\{titlesec\}.*?\n", "\n", tex)
    # amsmath
    if "\\usepackage{amsmath" not in tex:
        tex = tex.replace("\\begin{document}",
                          "\\usepackage{amsmath}\n\\begin{document}")
    return tex


def preprocess(tex_path: str) -> tuple[str, str]:
    """
    预处理 LaTeX 文件。

    Returns:
        (processed_tex_path, tmp_dir)
    """
    base_dir = os.path.dirname(os.path.abspath(tex_path))
    tmp_dir = tempfile.mkdtemp(prefix="texword_")

    print("[Phase 1] 预处理 LaTeX...")

    with open(tex_path, "r", encoding="utf-8") as f:
        tex = f.read()

    print("  展开 \\input{} 引用...")
    tex = resolve_inputs(tex, base_dir)
    tex = strip_comments(tex)

    print("  转换 PDF 图片...")
    tex = convert_pdf_figures(tex, base_dir, tmp_dir)

    print("  修复 pandoc 兼容性...")
    tex = fix_latex_for_pandoc(tex)

    processed_path = os.path.join(tmp_dir, "processed.tex")
    with open(processed_path, "w", encoding="utf-8") as f:
        f.write(tex)

    print(f"  预处理完成 → {processed_path}")
    return processed_path, tmp_dir
