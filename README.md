# TexWord

**学术论文格式转换器：LaTeX/PDF → 高质量可编辑 Word**

High-quality academic paper converter: **LaTeX/PDF → editable Word (OMML equations)**.

---

## 为什么需要 TexWord？ / Why TexWord?

学术论文从 LaTeX 转 Word 是科研工作者的常见痛点。现有工具要么公式变图片（不可编辑），要么格式严重丢失，要么只能输出 Markdown。

**TexWord 是第一个专为学术论文设计的 LaTeX → Word 转换器**，所有公式转为 Word 原生 OMML 格式（可编辑），同时保持期刊级排版质量。

---

## 与同类工具对比 / Comparison

| 工具 | 输入 | 输出 | 公式处理 | 学术结构 | 格式质量 |
|------|------|------|----------|----------|----------|
| **TexWord** | LaTeX, PDF | **Word (.docx)** | **OMML 原生可编辑** | **完整识别** | **期刊级** |
| pandoc | LaTeX, MD | Word, HTML... | OMML（但无后处理） | 不识别 | 基础 |
| pdf2docx | PDF | Word | 图片（不可编辑） | 不识别 | 中等 |
| MinerU (54k⭐) | PDF | Markdown | LaTeX 公式文本 | 部分识别 | 无排版 |
| marker | PDF | Markdown | LaTeX 公式文本 | 不识别 | 无排版 |
| nougat (Meta) | PDF | Markdown | LaTeX 公式文本 | 部分识别 | 无排版 |
| Mathpix | PDF, 图片 | LaTeX, MD | LaTeX 公式文本 | 部分识别 | 无排版 |

### 核心差异 / Key Differentiators

**1. 公式原生可编辑（OMML，非图片）**

pandoc 虽然也能输出 OMML，但直接使用 pandoc 转换学术论文会遇到大量兼容性问题（`threeparttable`、`fancyhdr`、`titlesec` 等包报错）。TexWord 的预处理层自动解决这些兼容性问题，确保所有公式正确转为可编辑 OMML。

**2. 学术论文专用后处理**

TexWord 理解学术论文结构（Title → Abstract → Sections → References），针对每个区域应用不同的格式规则：
- 标题：居中、16pt、加粗
- 摘要：缩进、11pt、1.5 倍行距
- 正文：首行缩进、12pt、双倍行距
- 参考文献：悬挂缩进（APA 格式）、11pt
- 图表标题：居中、10pt
- 表格：居中、学术三线表（booktabs 风格）、10pt

**3. 三阶段管线架构**

```
LaTeX → [预处理] → [pandoc 转换] → [python-docx 后处理] → Word
         ↓              ↓                    ↓
    修复兼容性      OMML 公式         字体/行距/缩进/
    展开 \input    保留结构           页眉/表格/参考文献
    PDF→PNG 图片
```

**4. 双路线输入（v2 规划）**

- LaTeX 路线：已完整可用
- PDF 路线：OCR 引擎抽象层已搭建（支持 DeepSeek-OCR-2），完整管线开发中

---

## 功能特性 / Features

- **公式 → OMML 原生可编辑**（在 Word 中双击即可编辑，非图片）
- **表格**完整转换，学术三线表（booktabs 风格）+ 表头加粗
- **PDF 图片**自动转 PNG（300 DPI）嵌入
- **完整格式控制**：Times New Roman、12pt、双倍行距、1 英寸边距
- **学术结构识别**：标题、摘要、各级标题、参考文献悬挂缩进
- **页眉**：运行标题 + 页码
- **寡行孤行控制**（widow/orphan control）
- **图表标题**居中格式化
- **自动清理** pandoc 转换残留

---

## 安装 / Install

```bash
git clone https://github.com/git-zyyang/texword.git
cd texword
pip install -e .
```

依赖：Python ≥ 3.9，pandoc

---

## 使用 / Usage

```bash
# LaTeX → Word（完整可用）
texword paper.tex
texword paper.tex -o output.docx
texword paper.tex --font-size 12 --line-spacing 2.0 --font "Times New Roman"

# PDF → Word（需要 OCR API key）
export DEEPSEEK_OCR_API_KEY="your-key"
texword paper.pdf
```

---

## 工作原理 / How It Works

**三阶段 LaTeX 管线：**

1. **预处理** — 展开 `\input{}`、PDF 图片转 PNG、修复 pandoc 不兼容的 LaTeX 包
2. **pandoc 转换** — LaTeX → DOCX，公式转为 OMML
3. **后处理** — python-docx 精修：字体、行距、标题层级、表格边框、参考文献缩进、图表标题、页面设置、寡行孤行控制

---

## 项目结构 / Structure

```
texword/
├── core/          # 转换编排 + 样式配置
├── latex/         # 预处理器、pandoc 桥接、后处理器
├── pdf/           # PDF 结构提取、OCR 引擎抽象层
│   └── engines/   # DeepSeek-OCR-2、通用 API
└── docx/          # (规划中) 直接 docx 构建器
```

---

## 路线图 / Roadmap

- [x] LaTeX → Word 完整管线
- [x] 公式 OMML 转换
- [x] 学术格式后处理（三线表、表头加粗、寡行孤行控制等）
- [x] PDF 路线框架 + OCR 引擎抽象
- [ ] PDF → Word 完整管线（OCR 识别 + 公式重建）
- [ ] 中文论文支持（宋体/黑体自动切换）
- [ ] 批量转换模式
- [ ] Web UI

---

## 支持作者

如果 TexWord 帮你省了时间，欢迎请作者喝杯咖啡 :coffee:

<img src="assets/donate_alipay.jpg" width="200" alt="支付宝赞赏码">

## License

MIT
