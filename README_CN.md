# doc2ai：面向 AI 工作流的文档转换插件

[English](README.md) / 中文

doc2ai 是一个用于 Claude Code 的文档转换插件，用于将 Office 文档转换为更适合 AI 读取和后续处理的文本格式。它重点保留源文档结构、清理格式噪音，让需求说明书、设计文档、表格数据等企业文档更容易被 AI Agent 和脚本分析。

## 安装

```bash
claude plugin marketplace add https://github.com/IronRookieCoder/doc2ai
claude plugin install doc2ai
```

## 使用方式

### Word 文档转 Markdown

```text
/doc2ai:docs2md input.docx
/doc2ai:docs2md input.doc -o md/
/doc2ai:docs2md docs/ --report
```

`docs2md` 技能将 `.doc` 和 `.docx` 文件转换为结构化 Markdown。它采用两阶段管道：

```text
doc/docx
  -> 脚本转换与清洗
  -> AI 靶向格式修复
  -> 最终 Markdown
```

### Excel 表格转 CSV

```text
/doc2ai:xlsx2csv report.xlsx
/doc2ai:xlsx2csv data/ -o csv/
```

`xlsx2csv` 技能将 `.xlsx` 文件转换为索引 CSV 和多个工作表 CSV。它保留原始网格结构，不做语义归一化。

## 技能

| 技能 | 命令 | 说明 |
| --- | --- | --- |
| `docs2md` | `/doc2ai:docs2md` | 将 `.doc` / `.docx` 文档转换为结构化 Markdown |
| `xlsx2csv` | `/doc2ai:xlsx2csv` | 将 `.xlsx` 工作簿转换为 AI 友好的 CSV 集合 |

## 依赖

### docs2md

- Pandoc 必须已安装并可在 `PATH` 中调用
- Python 3
- `pyyaml`
- 处理旧版 `.doc` 文件时，建议准备 WPS 或兼容的本地转换环境

### xlsx2csv

- Python 3
- `pandas`
- `python-calamine`
- `pyyaml`

缺少 Python 依赖时可按需安装：

```bash
pip install pandas python-calamine pyyaml
```

## 输出结构

### Markdown 输出

```text
md/
└── document.md
```

对于 `.doc` 输入，可能会在原文件旁生成并保留中间 `.docx` 文件。

使用 `--report` 时，转换报告输出到：

```text
md/
└── reports/
    └── document.json
```

### CSV 输出

```text
csv/
└── workbook/
    ├── workbook.csv
    ├── Sheet1.csv
    └── Sheet2.csv
```

工作簿同名 CSV 是索引文件，记录工作表顺序、工作表名、导出文件名和有效区域。

## 转换原则

- 保留源文档内容，不新增原文不存在的结论
- 优先做结构清洗，而不是还原视觉版式
- 保留电子表格原始网格，包括空白单元格、空白行和空白列
- 不推断表头，不做行级语义归一化
- 在明确安全时清理空锚点、图片残留、Pandoc 标注和异常表格格式等转换噪音
- 对可疑内容默认保留并交给人工复核，不主动删除

## 目录结构

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
└── xlsx2csv/
    ├── SKILL.md
    ├── config.yaml
    └── scripts/
```

## 注意事项

- 两个技能都支持目录输入。
- 批量转换会保留输入目录的相对层级，避免同名文件互相覆盖。
- 会跳过 `~$` 开头的 Office 临时文件。
- 内置脚本支持中文路径和中文文件名。
