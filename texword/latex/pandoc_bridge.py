"""pandoc 调用封装：LaTeX → docx 转换。"""

import os
import subprocess


def pandoc_convert(processed_tex: str, tmp_dir: str) -> str:
    """用 pandoc 将预处理后的 LaTeX 转为 docx。

    Returns:
        输出 docx 文件路径
    """
    print("\n[Phase 2] pandoc 转换...")

    output_path = os.path.join(tmp_dir, "raw_output.docx")

    cmd = [
        "pandoc",
        processed_tex,
        "-o", output_path,
        "-f", "latex",
        "-t", "docx",
        "--wrap=none",
        "--resource-path", tmp_dir,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0:
        print(f"  [WARN] pandoc warnings:\n{result.stderr[:500]}")

    if not os.path.exists(output_path):
        raise RuntimeError(f"pandoc failed: {result.stderr}")

    size_kb = os.path.getsize(output_path) / 1024
    print(f"  pandoc 转换完成 → {output_path} ({size_kb:.1f} KB)")
    return output_path
