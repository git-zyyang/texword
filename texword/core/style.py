"""StyleConfig — 文档样式配置"""


class StyleConfig:
    """文档样式配置，可通过命令行参数或代码覆盖。"""

    font_body: str = "Times New Roman"
    font_cjk: str = "宋体"
    font_size_body: int = 12       # pt
    font_size_title: int = 16
    font_size_h1: int = 14         # section
    font_size_h2: int = 13         # subsection
    font_size_h3: int = 12         # subsubsection
    font_size_abstract: int = 11
    font_size_table: int = 10
    font_size_note: int = 9
    font_size_ref: int = 11
    font_size_caption: int = 10    # 图表标题
    line_spacing: float = 2.0      # 双倍行距
    page_width: float = 21.59      # cm (Letter 8.5")
    page_height: float = 27.94     # cm (Letter 11")
    margin_top: float = 2.54       # cm
    margin_bottom: float = 2.54
    margin_left: float = 2.54
    margin_right: float = 2.54
    first_line_indent: float = 1.27  # cm (0.5 inch)
