---
name: md2ai
description: >
  将较长的 Markdown 文档整理为 AI 友好的渐进式披露结构。超过 500 行的长文档会被拆成
  主入口 TOC + 多个子文档，默认输出到 ai-native/，并生成 risk-index.json 供 AI 精准核实。
  当用户提到"优化 md 长文档""markdown 拆分""长文档 AI 读取""生成 ai-native""主入口 TOC"
  "渐进式披露""把大 markdown 拆成子文档"时触发。即使用户只是给出 .md 文件或目录并说明
  后续要交给 AI 阅读、审查、分析，也应使用此 skill。
---

# md2ai：长 Markdown → AI Native 文档结构

## 目标

把单个或批量 Markdown 文档整理成适合 AI 逐层读取的结构。主入口保持文档形态，只放标题和目录；运行元数据放在 `manifest.json`，风险核实入口放在 `risk-index.json`。

```text
ai-native/
└── 文档名/
    ├── 文档名.md          # 主入口，包含 TOC 和读取说明
    ├── 用户需求.md
    ├── 功能性需求.md
    ├── 非功能性需求.md
    ├── manifest.json      # 结构清单
    └── risk-index.json    # 需要 AI 精准核实的风险索引
```

核心原则：
- 不改写原文实质内容，只重组文件结构。
- 不在主入口混入源文件路径、行数、风险数量等过程信息。
- 前置片段如果只有文档标题和空行，不生成无实质内容的 `文档概述.md`。
- 长文档不直接读入上下文，优先运行脚本完成拆分。
- 有风险时只读取 `risk-index.json` 指向的子文档和行号做精准核实。
- 支持目录批量处理、嵌套子目录、中文目录和中文文件名。

## 调用参数

| 参数 | 说明 | 示例 |
| --- | --- | --- |
| `-o <dir>` / `--output-dir <dir>` | 指定输出目录，默认 `ai-native/` | `/md2ai docs/ -o ai-native/` |
| `--threshold <n>` | 超过多少行视为长文档，默认 `500` | `/md2ai a.md --threshold 800` |
| `--max-lines-per-doc <n>` | 拆分后单个子文档的目标最大行数，默认 `500` | `/md2ai a.md --max-lines-per-doc 400` |
| `--force` | 即使未超过阈值也生成主入口 + 子文档结构 | `/md2ai a.md --force` |

## 执行流程

### 步骤 1：确定输入

- 用户可能给出单个 `.md` 文件、目录或文件列表。
- 目录输入时递归扫描 `.md` 文件，保留相对目录层级。
- 跳过输出目录本身，避免重复处理已生成的 `ai-native/` 文件。

### 步骤 2：运行脚本拆分

默认输出到当前目录下的 `ai-native/`：

```bash
python <skill-path>/scripts/split_long_md.py <input.md|dir> -o ai-native/
```

如果用户指定输出目录，使用用户指定值：

```bash
python <skill-path>/scripts/split_long_md.py <input.md|dir> -o <用户指定目录>
```

脚本会：
1. 读取 Markdown 文件并统计行数。
2. 未超过阈值且未指定 `--force` 时，直接复制到输出目录并生成基础清单。
3. 超过阈值时，按 ATX 标题层级拆分为主入口和子文档。
4. 章节仍超过 `--max-lines-per-doc` 时，继续按下一级标题拆到子目录。
5. 无可用标题时，按固定行数切片，并把该情况写入风险索引。
6. 跳过只包含标题和空行的无效概述片段。
7. 生成 `manifest.json` 和 `risk-index.json`。

### 步骤 3：检查风险索引

脚本完成后，先读取每个输出文档目录下的 `risk-index.json`，不要直接读取全文。

风险索引为空：
- 向用户汇总输出路径、主入口文件和拆分数量即可。

风险索引非空：
- 按 `risks` 数组逐项处理。
- 每次只读取 `target` 指向的子文档。
- 如果风险项提供 `line`，围绕该行读取小范围上下文。
- 核实内容是否需要人工关注，必要时只在对应子文档中做最小修正。

典型风险：
- `no_heading`：原文缺少可拆分标题，只能按固定行数切片。
- `large_section_chunked`：某个大章节没有更深标题，只能切成若干部分。
- `unclosed_fence`：代码块围栏未闭合，可能影响后续标题识别。
- `html_table` / `grid_table`：复杂表格可能需要人工确认。
- `setext_heading`：疑似 Setext 标题，脚本不会把它当作拆分锚点。
- `heading_jump`：子文档内标题层级跳跃，可能来自源文档结构异常。
- `image_reference`：图片引用在拆分后可能仍需人工确认。

### 步骤 4：AI 精准核实协议

当存在风险时，按以下方式核实：

1. 读取 `risk-index.json`，只建立风险清单，不读全文。
2. 对每个风险项，打开 `target` 文件的相关局部。
3. 判断风险是否影响 AI 后续理解或内容完整性。
4. 能机械修复且不改变原文实质内容时，直接修复对应子文档。
5. 无法确定时，不猜测、不扩写，在最终汇总中列为需要人工关注。

核实时的输出格式：

```json
{
  "target": "功能性需求/接口需求.md",
  "risk_type": "heading_jump",
  "result": "needs_human_review",
  "note": "标题从 H3 跳到 H5，无法确认是否缺少中间标题"
}
```

## 输出结构

单文件输入：

```text
ai-native/
└── SFRD-REQM-03-5.4_XX项目XX模块系统需求规格说明书/
    ├── SFRD-REQM-03-5.4_XX项目XX模块系统需求规格说明书.md
    ├── 用户需求.md
    ├── 功能性需求.md
    ├── 非功能性需求.md
    ├── manifest.json
    └── risk-index.json
```

目录输入会保留相对层级：

```text
docs/
└── 需求/
    └── 模块A.md

ai-native/
└── 需求/
    └── 模块A/
        ├── 模块A.md
        ├── 用户需求.md
        ├── manifest.json
        └── risk-index.json
```

## 验收检查

处理完成后确认：
- 主入口文件存在，且包含 TOC。
- 主入口只包含文档标题和目录，不包含源文件路径、行数、风险数量等过程信息。
- 不生成只有标题和空行的 `文档概述.md` 或 `*_本节概述.md`。
- 超过阈值的文档被拆成文档目录，而不是单个超长文件。
- `manifest.json` 记录了主入口、子文档、行数和哈希。
- `risk-index.json` 存在；如 `risk_count > 0`，已经按风险索引做精准核实或汇总人工关注项。
- 中文路径、中文目录、中文文件名保持可读。
