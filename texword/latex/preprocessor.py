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


def _resolve_natbib_citations(tex: str) -> str:
    r"""将 natbib 引用命令（\citet, \citep, \citealt, \cite）解析为纯文本。

    pandoc 不能处理 natbib 命令 + thebibliography 的组合，会静默丢弃引用。
    此函数从 \bibitem[...]{key} 构建查找表，然后替换所有引用。

    支持两种 bibitem 格式：
      - \bibitem[Author, Year]{key}        （逗号分隔）
      - \bibitem[Author(Year)]{key}         （括号包裹年份）
      - \bibitem[Author and Author(Year)]{key}
      - \bibitem[Author et~al.(Year)]{key}
    """
    # 1. 构建 key → (author, year) 查找表
    bib_map = {}  # key → (author_str, year_str)
    for m in re.finditer(
        r'\\bibitem\[([^\]]+)\]\{([^}]+)\}', tex
    ):
        label, key = m.group(1).strip(), m.group(2).strip()
        # 尝试括号格式: "Author(Year)" 或 "Author et~al.(Year)"
        paren_match = re.match(r'^(.+?)\((\d{4}[a-z]?)\)$', label)
        if paren_match:
            author = paren_match.group(1).strip()
            year = paren_match.group(2).strip()
        elif "," in label:
            # 逗号格式: "Author, Year"
            author, year = label.rsplit(",", 1)
            author, year = author.strip(), year.strip()
        else:
            author, year = label, ""
        # 清理 et~al. → et al.
        author = author.replace("et~al.", "et al.").replace("et~al", "et al.")
        bib_map[key] = (author, year)

    if not bib_map:
        return tex

    def _lookup(key):
        key = key.strip()
        return bib_map.get(key, (key, ""))

    def _replace_citet(m):
        r"""\citet{key} → Author (Year)  |  \citet{k1, k2} → Author1 (Year1); Author2 (Year2)"""
        keys = m.group(1).split(",")
        parts = []
        for k in keys:
            author, year = _lookup(k)
            if year:
                parts.append(f"{author} ({year})")
            else:
                parts.append(author)
        return "; ".join(parts)

    def _replace_citep(m):
        r"""\citep{key} → (Author, Year)  |  \citep{k1, k2} → (Author1, Year1; Author2, Year2)"""
        keys = m.group(1).split(",")
        parts = []
        for k in keys:
            author, year = _lookup(k)
            if year:
                parts.append(f"{author}, {year}")
            else:
                parts.append(author)
        return "(" + "; ".join(parts) + ")"

    def _replace_citealt(m):
        r"""\citealt{key} → Author, Year（无括号）"""
        author, year = _lookup(m.group(1))
        if year:
            return f"{author}, {year}"
        return author

    count = 0
    # \citet 必须在 \cite 之前匹配（避免 \cite 吃掉 \citet 的前缀）
    tex, n = re.subn(r'\\citet\{([^}]+)\}', _replace_citet, tex)
    count += n
    tex, n = re.subn(r'\\citep\{([^}]+)\}', _replace_citep, tex)
    count += n
    tex, n = re.subn(r'\\citealt\{([^}]+)\}', _replace_citealt, tex)
    count += n
    # \cite{} 作为 fallback，当作 \citep 处理（括号引用）
    tex, n = re.subn(r'\\cite\{([^}]+)\}', _replace_citep, tex)
    count += n

    if count:
        print(f"    解析 {count} 个引用命令（citet/citep/citealt/cite）")
    return tex


def fix_latex_for_pandoc(tex: str) -> str:
    """修复 pandoc 不能正确处理的 LaTeX 结构。"""
    # \mathbb{数字} → \mathbf{数字}（amssymb 的 \mathbb 不支持数字，需要 bbm 包）
    tex = re.sub(r'\\mathbb\{(\d)\}', r'\\mathbf{\1}', tex)
    # \boldsymbol{x} → \mathbf{x}（pandoc OMML 不支持 \boldsymbol）
    tex = re.sub(r'\\boldsymbol\{([^}]+)\}', r'\\mathbf{\1}', tex)
    # {\em ...} → \emph{...}（pandoc 不支持旧式 {\em} 语法，会丢弃内容）
    tex = re.sub(r'\{\\em\s+([^}]+)\}', r'\\emph{\1}', tex)
    # {\bf ...} → \textbf{...}
    tex = re.sub(r'\{\\bf\s+([^}]+)\}', r'\\textbf{\1}', tex)
    # {\it ...} → \textit{...}
    tex = re.sub(r'\{\\it\s+([^}]+)\}', r'\\textit{\1}', tex)
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
    # natbib 引用命令 → 纯文本（pandoc 不支持 natbib + thebibliography）
    tex = _resolve_natbib_citations(tex)
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
