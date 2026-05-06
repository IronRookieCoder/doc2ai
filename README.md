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
/doc2ai:docs2md input.docx --config custom.yaml
```

The `docs2md` skill converts `.doc` and `.docx` files into structured Markdown. It uses a two-stage pipeline:

```text
doc/docx
  -> script conversion and cleanup (Pandoc + Lua filter + regex)
  -> targeted AI formatting repair (only risk-flagged regions)
  -> final Markdown
```

| Parameter | Description | Default |
| --- | --- | --- |
| `-o <dir>` | Markdown output directory | `md/` |
| `--config <path>` | Path to configuration file | `config.yaml` in the skill directory |
| `--report` | Generate a JSON conversion report | not generated |

### Convert Long Markdown to AI Native Structure

```text
/doc2ai:md2ai input.md
/doc2ai:md2ai md/ -o ai-native/
/doc2ai:md2ai input.md --threshold 800
/doc2ai:md2ai input.md --force
```

The `md2ai` skill splits Markdown files longer than the configured threshold (default 500 lines) into a main TOC entry plus focused child documents. Output goes to the current directory by default. During processing it uses a risk index so AI verifies only risky child documents; process JSON files are removed before final delivery.

| Parameter | Description | Default |
| --- | --- | --- |
| `-o <dir>` | Output directory | current directory `.` |
| `--threshold <n>` | Line count above which a file is treated as long | `500` |
| `--max-lines-per-doc <n>` | Target maximum lines per child document | `500` |
| `--force` | Generate TOC + child structure even below the threshold | off |
| `--keep-process-files` | Keep `manifest.json`, `risk-index.json`, and `summary.json` | not kept |

### Convert Spreadsheets to CSV

```text
/doc2ai:xlsx2csv report.xlsx
/doc2ai:xlsx2csv data/ -o csv/
```

The `xlsx2csv` skill converts `.xlsx` files into an index CSV plus one CSV file per worksheet. It preserves the original grid layout and avoids semantic normalization.

| Parameter | Description | Default |
| --- | --- | --- |
| `-o <dir>` | CSV output directory | `csv/` |

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

Output goes to the current directory by default, creating a subdirectory named after the document:

```text
./
└── document/
    ├── document.md
    ├── User requirements.md
    ├── Functional requirements.md
    └── Non-functional requirements.md
```

When `-o ai-native/` is specified:

```text
ai-native/
└── document/
    ├── document.md
    ├── User requirements.md
    ├── Functional requirements.md
    └── Non-functional requirements.md
```

Batch processing preserves relative input subdirectories. `manifest.json`, `risk-index.json`, and `summary.json` are process files used during verification and removed before final delivery.

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
- `xlsx2csv` skips hidden worksheets by default. Set `sheet.include_hidden_sheets: true` in `config.yaml` to export them.
- When converting `.doc` files, `docs2md` skips the `.doc` if a matching `.docx` already exists in the same directory to avoid duplicate outputs.
