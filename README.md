# doc2ai: Document Conversion Plugin for AI Workflows

English / [中文](README_CN.md)

doc2ai is a Claude Code plugin for converting office documents into AI-friendly text formats. It focuses on preserving source structure while removing format noise, so downstream AI agents and scripts can inspect requirements, designs, spreadsheets, and other enterprise documents more reliably.

## Installation

```bash
claude plugin marketplace add https://github.com/IronRookieCoder/doc2ai
claude plugin install doc2ai
```

## Usage

### Convert Word Documents to Markdown

```text
/doc2ai:docs2md input.docx
/doc2ai:docs2md input.doc -o md/
/doc2ai:docs2md docs/ --report
```

The `docs2md` skill converts `.doc` and `.docx` files into structured Markdown. It uses a two-stage pipeline:

```text
doc/docx
  -> script conversion and cleanup
  -> targeted AI formatting repair
  -> final Markdown
```

### Convert Long Markdown to AI Native Structure

```text
/doc2ai:md2ai input.md
/doc2ai:md2ai md/ -o ai-native/
```

The `md2ai` skill splits Markdown files longer than 500 lines into a main TOC entry plus focused child documents under `ai-native/`. It also writes `risk-index.json` so AI can verify only risky child documents and local line ranges instead of loading the whole file into context.

### Convert Spreadsheets to CSV

```text
/doc2ai:xlsx2csv report.xlsx
/doc2ai:xlsx2csv data/ -o csv/
```

The `xlsx2csv` skill converts `.xlsx` files into an index CSV plus one CSV file per worksheet. It preserves the original grid layout and avoids semantic normalization.

## Skills

| Skill | Command | Description |
| --- | --- | --- |
| `docs2md` | `/doc2ai:docs2md` | Convert `.doc` / `.docx` documents into structured Markdown |
| `md2ai` | `/doc2ai:md2ai` | Split long Markdown into an AI Native TOC and child documents |
| `xlsx2csv` | `/doc2ai:xlsx2csv` | Convert `.xlsx` workbooks into AI-friendly CSV collections |

## Dependencies

### docs2md

- Pandoc must be installed and available in `PATH`
- Python 3
- `pyyaml`
- WPS or a compatible local conversion environment is recommended for legacy `.doc` files

### xlsx2csv

- Python 3
- `pandas`
- `python-calamine`
- `pyyaml`

Install missing Python dependencies when needed:

```bash
pip install pandas python-calamine pyyaml
```

## Output Structure

### Markdown Output

```text
md/
└── document.md
```

For `.doc` inputs, an intermediate `.docx` file may be generated and retained beside the original file.

When `--report` is used, conversion reports are written under:

```text
md/
└── reports/
    └── document.json
```

### AI Native Markdown Output

```text
ai-native/
└── document/
    ├── document.md
    ├── User requirements.md
    ├── Functional requirements.md
    ├── manifest.json
    └── risk-index.json
```

Batch processing preserves relative input subdirectories. `risk-index.json` tells AI which child documents and local line ranges need focused verification.

### CSV Output

```text
csv/
└── workbook/
    ├── workbook.csv
    ├── Sheet1.csv
    └── Sheet2.csv
```

The workbook-level CSV is an index file that records worksheet order, worksheet name, exported file name, and used range.

## Conversion Principles

- Preserve source content and avoid adding conclusions not present in the original file
- Prefer structural cleanup over visual layout restoration
- Keep original spreadsheet grids, including blank cells, blank rows, and blank columns
- Do not infer spreadsheet headers or normalize rows
- Remove conversion noise such as empty anchors, image remnants, Pandoc annotations, and invalid table formatting when clearly safe
- Keep suspicious content for human review instead of deleting it by default

## Directory Structure

```text
.claude-plugin/
├── plugin.json
└── marketplace.json
skills/
├── docs2md/
│   ├── SKILL.md
│   ├── config.yaml
│   ├── scripts/
│   └── references/
├── md2ai/
│   ├── SKILL.md
│   └── scripts/
└── xlsx2csv/
    ├── SKILL.md
    ├── config.yaml
    └── scripts/
```

## Notes

- Directory input is supported for all three skills.
- Batch conversion preserves relative subdirectories to avoid filename collisions.
- Office temporary files starting with `~$` are skipped.
- Chinese paths and filenames are supported by the bundled scripts.
