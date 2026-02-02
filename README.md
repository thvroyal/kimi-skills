# Kimi Skills

A collection of AI agent skills extracted from [Kimi](https://www.kimi.com) that enable advanced document generation capabilities.

> **Attribution**: These skills were originally developed by [Moonshot AI](https://www.moonshot.cn/) for their Kimi AI assistant.

## Why Kimi Skills?

Compared to Anthropic document skills, Kimi Skills offer several enhancements:

| Aspect | Kimi Skills | Anthropic Skills |
|--------|-------------|-----------------|
| **Visual Quality** | Designer-quality covers, Morandi/ink-wash backgrounds, professional themes | Basic formatting |
| **Document Structure** | Auto-generates TOC, headers/footers, cover & back cover pages | Manual structure |
| **Charts** | Native Word/Excel charts (editable, small file size) | Image-based charts |
| **Validation** | OpenXML validators, formula checkers, reference verification | Basic output |
| **Format Support** | Full OpenXML manipulation, track changes, comments | Limited editing |
| **Academic Features** | Citations (GB/T 7714, APA), cross-references, math equations | Basic citations |
| **Excel** | PivotTables, currency formatting, cover pages with metrics | Data only |

**Key advantages:**
- 🎨 **Beautiful output** - Professional designer-quality documents, not generic templates
- 🔧 **Deep control** - Direct XML manipulation for precise formatting
- ✅ **Validation pipeline** - Multiple validation steps ensure error-free output
- 📊 **Native charts** - Editable charts instead of static images

## Installation

```bash
npx skills add thvroyal/kimi-skills
```

## Skills Overview

| Skill | Description |
|-------|-------------|
| **docx** | Generate and edit Word documents (.docx) with professional layouts, charts, track-changes, and more |
| **pdf** | Create PDFs using HTML+Paged.js or LaTeX. Process existing PDFs (extract, merge, fill forms) |
| **xlsx** | Advanced spreadsheet manipulation with formulas, formatting, charts, and PivotTables |

## Skills

### 📄 DOCX - Word Document Generation

Creates professional Word documents using C# OpenXML SDK for new documents and Python+lxml for editing existing ones.

**Features:**
- Professional covers and back covers with designer-quality backgrounds
- Native Word charts (pie, bar, line)
- Table of Contents with refresh hints
- Headers/footers with page numbers
- Comments and Track Changes support
- Math equations (OMML)
- Morandi and ink-wash style backgrounds

### 📑 PDF - Professional PDF Solution

Two routes for PDF creation:
- **HTML Route** (default): Uses Playwright + Paged.js for HTML→PDF conversion
- **LaTeX Route**: Uses Tectonic for LaTeX compilation

**Features:**
- Academic paper formatting (IMRaD structure)
- KaTeX math formulas
- Mermaid diagrams
- Three-line tables
- Citations (GB/T 7714 for Chinese, APA for English)
- 11 cover style options
- Process existing PDFs (extract, merge, split, fill forms)

### 📊 XLSX - Excel Spreadsheet Manipulation

Creates and manipulates Excel files using Python + openpyxl/pandas.

**Features:**
- Complex formulas with automatic validation
- Professional styling (monochrome/finance themes)
- Native Excel charts (bar, line, pie, area)
- PivotTables with charts
- Cover pages with key metrics summary
- Currency formatting for financial data
- Source citation tracking for external data

## Examples

The `examples/` directory contains sample outputs:

- `faker-career-analysis.pdf` - PDF report example
- `faker-champion-stats.xlsx` - Excel spreadsheet with charts
- `generate_faker_docx.py` - Python script for DOCX generation

## Structure

```
kimi-skills/
├── skills/
│   ├── docx/             # Word document skill
│   │   ├── SKILL.md
│   │   ├── scripts/
│   │   ├── references/
│   │   ├── assets/
│   │   └── validator/
│   ├── pdf/              # PDF skill
│   │   ├── SKILL.md
│   │   ├── routes/
│   │   └── scripts/
│   └── xlsx/             # Excel skill
│       ├── SKILL.md
│       ├── scripts/
│       └── pivot-table.md
└── examples/
```

## Requirements

Skills automatically detect and install dependencies:

- **docx**: .NET SDK, Python 3, pandoc (optional)
- **pdf**: Node.js, Playwright, Python 3 (for processing)
- **xlsx**: Python 3, openpyxl, pandas

## Credits

- **Original Skills**: [Moonshot AI](https://www.moonshot.cn/) - Kimi AI Assistant
- **Packaging**: This repository

## License

These skills are extracted from Kimi for educational and personal use. All rights to the original implementation belong to Moonshot AI.
