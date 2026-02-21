# TexWord

High-quality academic paper converter: **LaTeX/PDF → editable Word**.

Unlike generic converters (pandoc alone, pdf2docx), TexWord is purpose-built for academic papers — equations become native editable OMML, tables keep their structure, and formatting follows journal standards.

## Features

- **105 equations → native OMML** (editable in Word, not images)
- **12 tables** with borders and centered formatting
- **6 PDF figures** auto-converted to PNG at 300 DPI
- **Full format control**: Times New Roman, 12pt, double spacing, 1-inch margins
- **Academic structure**: title, abstract, headings, references with hanging indent
- **Page headers** with running title + page numbers
- **Widow/orphan control** for clean page breaks
- **Figure/table captions** centered, 10pt

## Install

```bash
pip install -e .
```

Requires: Python ≥ 3.9, pandoc

## Usage

```bash
# LaTeX → Word (fully functional)
texword paper.tex
texword paper.tex -o output.docx
texword paper.tex --font-size 12 --line-spacing 2.0

# PDF → Word (framework, requires OCR API key)
export DEEPSEEK_OCR_API_KEY="your-key"
texword paper.pdf
```

## How It Works

**3-phase LaTeX pipeline:**

1. **Preprocess** — expand `\input{}`, convert PDF figures to PNG, fix pandoc compatibility
2. **Pandoc convert** — LaTeX → DOCX with OMML math
3. **Post-process** — python-docx fixes: fonts, spacing, headings, tables, references, captions, page setup

## Project Structure

```
texword/
├── core/          # converter orchestration + style config
├── latex/         # preprocessor, pandoc bridge, postprocessor
├── pdf/           # PDF extractor, OCR engine abstraction
│   └── engines/   # DeepSeek-OCR-2, generic API
└── docx/          # (future) direct docx builders
```

## License

MIT
