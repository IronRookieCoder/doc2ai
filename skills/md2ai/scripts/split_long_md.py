#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将长 Markdown 文档拆分为 AI 友好的主入口 + 子文档结构。

用法：
    python split_long_md.py <input.md|dir> -o ai-native/
    python split_long_md.py docs/ -o ai-native/ --threshold 500 --max-lines-per-doc 500
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path


HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
FENCE_RE = re.compile(r"^\s*(```+|~~~+)")
SETEXT_RE = re.compile(r"^\s*(=+|-{3,})\s*$")
HTML_TABLE_RE = re.compile(r"<\s*/?\s*table\b", re.IGNORECASE)
GRID_TABLE_RE = re.compile(r"^\s*\+[-=+ ]+\+\s*$")
IMAGE_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")
INVALID_NAME_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
SPACE_RE = re.compile(r"\s+")


@dataclass
class Heading:
    index: int
    level: int
    title: str


@dataclass
class OutputDoc:
    title: str
    path: Path
    line_count: int
    sha256: str
    source_start_line: int | None = None
    source_end_line: int | None = None


class NameAllocator:
    def __init__(self) -> None:
        self._used: dict[str, set[str]] = {}

    def unique(self, directory: Path, stem: str, suffix: str = "") -> str:
        key = os.path.normcase(str(directory.resolve()))
        used = self._used.setdefault(key, set())
        safe_stem = sanitize_name(stem) or "未命名"
        candidate = f"{safe_stem}{suffix}"
        base = safe_stem
        seq = 2
        while candidate.casefold() in used:
            candidate = f"{base}_{seq}{suffix}"
            seq += 1
        used.add(candidate.casefold())
        return candidate


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def resolve_path(raw_path: str) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    return path.resolve()


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def read_markdown(path: Path) -> list[str]:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return data.decode(encoding).splitlines(keepends=True)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace").splitlines(keepends=True)


def write_lines(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("".join(lines), encoding="utf-8", newline="")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def sha256_lines(lines: list[str]) -> str:
    return hashlib.sha256("".join(lines).encode("utf-8")).hexdigest()


def sanitize_name(name: str, max_length: int = 80) -> str:
    name = INVALID_NAME_RE.sub("_", name)
    name = SPACE_RE.sub("_", name.strip())
    name = name.strip(" ._")
    if len(name) > max_length:
        name = name[:max_length].rstrip(" ._")
    return name or "未命名"


def clean_heading_title(title: str) -> str:
    title = re.sub(r"\s+\{#[^}]+\}\s*$", "", title.strip())
    title = re.sub(r"\s+", " ", title)
    return title.strip("# ").strip() or "未命名"


def has_meaningful_body(lines: list[str]) -> bool:
    """判断前置片段是否包含标题之外的正文内容。"""
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if HEADING_RE.match(stripped):
            continue
        return True
    return False


def collect_headings(lines: list[str]) -> list[Heading]:
    headings: list[Heading] = []
    fence_marker: str | None = None

    for idx, line in enumerate(lines):
        fence_match = FENCE_RE.match(line)
        if fence_match:
            marker = fence_match.group(1)
            marker_char = marker[0]
            if fence_marker is None:
                fence_marker = marker
            elif marker_char == fence_marker[0] and len(marker) >= len(fence_marker):
                fence_marker = None
            continue

        if fence_marker is not None:
            continue

        match = HEADING_RE.match(line.rstrip("\n\r"))
        if match:
            headings.append(
                Heading(
                    index=idx,
                    level=len(match.group(1)),
                    title=clean_heading_title(match.group(2)),
                )
            )

    return headings


def has_unclosed_fence(lines: list[str]) -> bool:
    fence_marker: str | None = None
    for line in lines:
        match = FENCE_RE.match(line)
        if not match:
            continue
        marker = match.group(1)
        if fence_marker is None:
            fence_marker = marker
        elif marker[0] == fence_marker[0] and len(marker) >= len(fence_marker):
            fence_marker = None
    return fence_marker is not None


def relative_link(path: Path) -> str:
    return path.as_posix()


def choose_split_level(headings: list[Heading]) -> int | None:
    if not headings:
        return None
    min_level = min(h.level for h in headings)
    if headings[0].index <= 2 and headings[0].level == 1 and any(h.level == 2 for h in headings):
        return 2
    return min_level


def split_by_heading_level(lines: list[str], level: int) -> list[tuple[Heading, int, int]]:
    headings = [h for h in collect_headings(lines) if h.level == level]
    sections: list[tuple[Heading, int, int]] = []
    for idx, heading in enumerate(headings):
        end = headings[idx + 1].index if idx + 1 < len(headings) else len(lines)
        sections.append((heading, heading.index, end))
    return sections


def detect_leaf_risks(
    lines: list[str],
    rel_path: Path,
    risks: list[dict],
    source_file: Path,
    source_start_line: int | None,
) -> None:
    rel = relative_link(rel_path)

    def add(risk_type: str, line_no: int | None, reason: str, suggestion: str, severity: str = "medium") -> None:
        risks.append(
            {
                "type": risk_type,
                "severity": severity,
                "target": rel,
                "line": line_no,
                "source": str(source_file),
                "source_line": None if source_start_line is None or line_no is None else source_start_line + line_no - 1,
                "reason": reason,
                "suggested_check": suggestion,
            }
        )

    if has_unclosed_fence(lines):
        add(
            "unclosed_fence",
            None,
            "检测到未闭合的代码块围栏，可能影响标题识别或后续渲染。",
            "读取该子文档，确认代码块边界是否来自原文，必要时只修复围栏。",
            "high",
        )

    headings = collect_headings(lines)
    previous_level: int | None = None
    for heading in headings:
        if previous_level is not None and heading.level > previous_level + 1:
            add(
                "heading_jump",
                heading.index + 1,
                f"标题层级从 H{previous_level} 跳到 H{heading.level}。",
                "围绕该标题读取上下文，确认是否存在标题缺失或转换造成的层级错误。",
            )
        previous_level = heading.level

    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()
        if HTML_TABLE_RE.search(line):
            add(
                "html_table",
                idx,
                "检测到 HTML table，可能是复杂表格转换残留。",
                "检查该表格是否完整、是否需要保留 HTML 或转换为 GFM 表格。",
            )
        elif GRID_TABLE_RE.match(line):
            add(
                "grid_table",
                idx,
                "检测到 grid table 边框，可能不适合后续 AI 直接阅读。",
                "读取完整表格区域，确认是否需要转换为 GFM pipe table。",
            )
        elif IMAGE_RE.search(line):
            add(
                "image_reference",
                idx,
                "检测到图片引用，拆分后可能仍需人工确认图片内容。",
                "确认图片是否是关键信息；如无法读取图片，在最终汇总中提示人工关注。",
                "low",
            )
        elif len(stripped) > 1200:
            add(
                "very_long_line",
                idx,
                "检测到超长行，可能是表格、粘连段落或转换噪音。",
                "读取该行附近上下文，确认是否需要拆行或保留原样。",
            )

        if idx > 1 and SETEXT_RE.match(line):
            previous = lines[idx - 2].strip()
            if previous and not previous.startswith("|"):
                add(
                    "setext_heading",
                    idx,
                    "检测到疑似 Setext 标题，脚本不会将其作为拆分锚点。",
                    "确认该位置是否应视为标题；如影响结构，可局部改为 ATX 标题。",
                    "low",
                )


def add_structural_risk(
    risks: list[dict],
    risk_type: str,
    target: Path,
    reason: str,
    suggestion: str,
    source_file: Path,
    severity: str = "medium",
) -> None:
    risks.append(
        {
            "type": risk_type,
            "severity": severity,
            "target": relative_link(target),
            "line": None,
            "source": str(source_file),
            "source_line": None,
            "reason": reason,
            "suggested_check": suggestion,
        }
    )


def make_source_ranges(base_start_line: int | None, start: int, end: int) -> tuple[int | None, int | None]:
    if base_start_line is None:
        return None, None
    return base_start_line + start, base_start_line + end - 1


def emit_leaf(
    title: str,
    lines: list[str],
    output_path: Path,
    rel_path: Path,
    docs: list[OutputDoc],
    risks: list[dict],
    source_file: Path,
    source_start_line: int | None,
    source_end_line: int | None,
) -> None:
    write_lines(output_path, lines)
    docs.append(
        OutputDoc(
            title=title,
            path=rel_path,
            line_count=len(lines),
            sha256=sha256_lines(lines),
            source_start_line=source_start_line,
            source_end_line=source_end_line,
        )
    )
    detect_leaf_risks(lines, rel_path, risks, source_file, source_start_line)


def emit_chunked_section(
    title: str,
    lines: list[str],
    base_dir: Path,
    rel_dir: Path,
    allocator: NameAllocator,
    docs: list[OutputDoc],
    risks: list[dict],
    source_file: Path,
    source_start_line: int | None,
    max_lines: int,
) -> list[dict]:
    entries: list[dict] = []
    for part_index, start in enumerate(range(0, len(lines), max_lines), start=1):
        end = min(start + max_lines, len(lines))
        part_title = f"{title}_第{part_index}部分"
        filename = allocator.unique(base_dir, part_title, ".md")
        out_path = base_dir / filename
        rel_path = rel_dir / filename
        chunk_lines = lines[start:end]
        src_start, src_end = make_source_ranges(source_start_line, start, end)
        emit_leaf(part_title, chunk_lines, out_path, rel_path, docs, risks, source_file, src_start, src_end)
        add_structural_risk(
            risks,
            "large_section_chunked",
            rel_path,
            "章节超过目标行数且没有更深层标题，已按固定行数切片。",
            "抽查相邻分片的边界，确认句子、表格或列表没有被切断到影响理解。",
            source_file,
        )
        entries.append({"title": part_title, "path": rel_path, "line_count": len(chunk_lines)})
    return entries


def emit_section(
    title: str,
    lines: list[str],
    base_dir: Path,
    rel_dir: Path,
    allocator: NameAllocator,
    docs: list[OutputDoc],
    risks: list[dict],
    source_file: Path,
    source_start_line: int | None,
    max_lines: int,
    current_level: int | None,
) -> dict:
    if len(lines) <= max_lines:
        filename = allocator.unique(base_dir, title, ".md")
        out_path = base_dir / filename
        rel_path = rel_dir / filename
        source_end_line = None if source_start_line is None else source_start_line + len(lines) - 1
        emit_leaf(title, lines, out_path, rel_path, docs, risks, source_file, source_start_line, source_end_line)
        return {"title": title, "path": rel_path, "line_count": len(lines), "children": []}

    headings = collect_headings(lines)
    deeper = [h for h in headings if current_level is None or h.level > current_level]
    if not deeper:
        return {
            "title": title,
            "path": None,
            "line_count": len(lines),
            "children": emit_chunked_section(
                title,
                lines,
                base_dir,
                rel_dir,
                allocator,
                docs,
                risks,
                source_file,
                source_start_line,
                max_lines,
            ),
        }

    next_level = min(h.level for h in deeper)
    child_sections = split_by_heading_level(lines, next_level)
    if not child_sections:
        return {
            "title": title,
            "path": None,
            "line_count": len(lines),
            "children": emit_chunked_section(
                title,
                lines,
                base_dir,
                rel_dir,
                allocator,
                docs,
                risks,
                source_file,
                source_start_line,
                max_lines,
            ),
        }

    dir_name = allocator.unique(base_dir, title)
    section_dir = base_dir / dir_name
    section_rel_dir = rel_dir / dir_name
    section_dir.mkdir(parents=True, exist_ok=True)

    entry_filename = allocator.unique(section_dir, title, ".md")
    entry_rel_path = section_rel_dir / entry_filename
    children: list[dict] = []

    if current_level is not None and next_level > current_level + 1:
        add_structural_risk(
            risks,
            "heading_jump",
            entry_rel_path,
            f"章节内下一级标题从 H{current_level} 跳到 H{next_level}。",
            "读取该章节入口及其下级目录，确认是否缺少中间标题或源文档层级异常。",
            source_file,
        )

    first_child_start = child_sections[0][1]
    if first_child_start > 0 and has_meaningful_body(lines[:first_child_start]):
        overview_title = f"{title}_本节概述"
        overview_lines = lines[:first_child_start]
        src_start, _ = make_source_ranges(source_start_line, 0, first_child_start)
        overview = emit_section(
            overview_title,
            overview_lines,
            section_dir,
            section_rel_dir,
            allocator,
            docs,
            risks,
            source_file,
            src_start,
            max_lines,
            current_level,
        )
        children.append(overview)

    for child_heading, start, end in child_sections:
        child_start_line, _ = make_source_ranges(source_start_line, start, end)
        child = emit_section(
            child_heading.title,
            lines[start:end],
            section_dir,
            section_rel_dir,
            allocator,
            docs,
            risks,
            source_file,
            child_start_line,
            max_lines,
            child_heading.level,
        )
        children.append(child)

    entry_text = render_section_entry(title, children, entry_rel_path)
    write_text(section_dir / entry_filename, entry_text)
    docs.append(
        OutputDoc(
            title=title,
            path=entry_rel_path,
            line_count=len(entry_text.splitlines()),
            sha256=hashlib.sha256(entry_text.encode("utf-8")).hexdigest(),
            source_start_line=None,
            source_end_line=None,
        )
    )
    return {"title": title, "path": entry_rel_path, "line_count": len(lines), "children": children}


def flatten_entries(entries: list[dict]) -> list[dict]:
    flat: list[dict] = []
    for entry in entries:
        if entry.get("path") is not None:
            flat.append(entry)
        flat.extend(flatten_entries(entry.get("children", [])))
    return flat


def render_toc_entries(entries: list[dict], base_path: Path, depth: int = 0) -> list[str]:
    lines: list[str] = []
    indent = "  " * depth
    for entry in entries:
        path = entry.get("path")
        line_count = entry.get("line_count", 0)
        if path is not None:
            target = os.path.relpath(path, start=base_path.parent).replace("\\", "/")
            lines.append(f"{indent}- [{entry['title']}]({target})（{line_count} 行）\n")
        else:
            lines.append(f"{indent}- {entry['title']}（{line_count} 行）\n")
        children = entry.get("children", [])
        if children:
            lines.extend(render_toc_entries(children, base_path, depth + 1))
    return lines


def render_main_entry(title: str, entries: list[dict]) -> str:
    lines = [
        f"# {title}\n",
        "\n",
        "## 目录\n",
        "\n",
    ]
    lines.extend(render_toc_entries(entries, Path(f"{title}.md")))
    return "".join(lines)


def render_section_entry(title: str, children: list[dict], entry_path: Path) -> str:
    lines = [
        f"# {title}\n",
        "\n",
        "## 目录\n",
        "\n",
    ]
    lines.extend(render_toc_entries(children, entry_path))
    return "".join(lines)


def write_manifest(doc_dir: Path, source: Path, source_lines: int, main_file: Path, docs: list[OutputDoc], risks: list[dict], mode: str) -> None:
    manifest = {
        "version": 1,
        "mode": mode,
        "source": str(source),
        "source_lines": source_lines,
        "main": relative_link(main_file),
        "document_count": len(docs),
        "risk_count": len(risks),
        "documents": [
            {
                "title": doc.title,
                "path": relative_link(doc.path),
                "line_count": doc.line_count,
                "sha256": doc.sha256,
                "source_start_line": doc.source_start_line,
                "source_end_line": doc.source_end_line,
            }
            for doc in docs
        ],
    }
    write_text(doc_dir / "manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))


def write_risk_index(doc_dir: Path, source: Path, risks: list[dict]) -> None:
    risk_index = {
        "version": 1,
        "source": str(source),
        "risk_count": len(risks),
        "usage": "AI 只读取本索引中 target 指向的子文档和相关行号做精准核实，不要直接读取全文。",
        "risks": risks,
    }
    write_text(doc_dir / "risk-index.json", json.dumps(risk_index, ensure_ascii=False, indent=2))


def copy_short_document(
    input_path: Path,
    output_root: Path,
    relative_source: Path,
    lines: list[str],
    risks: list[dict],
) -> Path:
    output_path = output_root / relative_source
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(input_path, output_path)
    rel_path = Path(relative_source.name)
    detect_leaf_risks(lines, rel_path, risks, input_path, 1)
    docs = [
        OutputDoc(
            title=input_path.stem,
            path=rel_path,
            line_count=len(lines),
            sha256=sha256_lines(lines),
            source_start_line=1,
            source_end_line=len(lines),
        )
    ]
    write_manifest(output_path.parent, input_path, len(lines), Path(relative_source.name), docs, risks, "copy")
    write_risk_index(output_path.parent, input_path, risks)
    return output_path


def split_long_document(
    input_path: Path,
    output_root: Path,
    relative_source: Path,
    lines: list[str],
    max_lines: int,
) -> Path:
    allocator = NameAllocator()
    doc_title = sanitize_name(input_path.stem)
    doc_dir = output_root / relative_source.parent / doc_title
    if doc_dir.exists():
        shutil.rmtree(doc_dir)
    doc_dir.mkdir(parents=True, exist_ok=True)

    main_filename = allocator.unique(doc_dir, doc_title, ".md")
    main_rel_path = Path(main_filename)
    main_path = doc_dir / main_filename
    docs: list[OutputDoc] = []
    risks: list[dict] = []
    entries: list[dict] = []

    headings = collect_headings(lines)
    split_level = choose_split_level(headings)

    if split_level is None:
        chunk_entries = emit_chunked_section(
            doc_title,
            lines,
            doc_dir,
            Path("."),
            allocator,
            docs,
            risks,
            input_path,
            1,
            max_lines,
        )
        entries.extend(chunk_entries)
        add_structural_risk(
            risks,
            "no_heading",
            chunk_entries[0]["path"] if chunk_entries else main_rel_path,
            "原文没有可识别的 ATX 标题，只能按固定行数切片。",
            "抽查每个分片边界，确认段落、列表和表格没有被切断到影响理解。",
            input_path,
            "high",
        )
    else:
        sections = split_by_heading_level(lines, split_level)
        first_start = sections[0][1] if sections else 0
        if first_start > 0 and has_meaningful_body(lines[:first_start]):
            overview_title = "文档概述"
            overview = emit_section(
                overview_title,
                lines[:first_start],
                doc_dir,
                Path("."),
                allocator,
                docs,
                risks,
                input_path,
                1,
                max_lines,
                None,
            )
            entries.append(overview)

        for heading, start, end in sections:
            entry = emit_section(
                heading.title,
                lines[start:end],
                doc_dir,
                Path("."),
                allocator,
                docs,
                risks,
                input_path,
                start + 1,
                max_lines,
                heading.level,
            )
            entries.append(entry)

    main_text = render_main_entry(doc_title, entries)
    write_text(main_path, main_text)
    docs.insert(
        0,
        OutputDoc(
            title=doc_title,
            path=main_rel_path,
            line_count=len(main_text.splitlines()),
            sha256=hashlib.sha256(main_text.encode("utf-8")).hexdigest(),
            source_start_line=None,
            source_end_line=None,
        ),
    )
    write_manifest(doc_dir, input_path, len(lines), main_rel_path, docs, risks, "split")
    write_risk_index(doc_dir, input_path, risks)
    return main_path


def discover_markdown_files(input_path: Path, output_root: Path) -> tuple[list[Path], Path | None]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".md":
            return [], None
        return [input_path], None

    if not input_path.is_dir():
        return [], None

    files = []
    for path in sorted(input_path.rglob("*.md"), key=lambda p: str(p.relative_to(input_path)).casefold()):
        if path.name.startswith("."):
            continue
        if is_relative_to(path, output_root):
            continue
        files.append(path.resolve())
    return files, input_path


def process_file(
    input_path: Path,
    output_root: Path,
    input_root: Path | None,
    threshold: int,
    max_lines: int,
    force: bool,
) -> dict:
    relative_source = Path(input_path.name) if input_root is None else input_path.relative_to(input_root)
    lines = read_markdown(input_path)
    risks: list[dict] = []

    if len(lines) <= threshold and not force:
        output_path = copy_short_document(input_path, output_root, relative_source, lines, risks)
        mode = "copy"
    else:
        output_path = split_long_document(input_path, output_root, relative_source, lines, max_lines)
        mode = "split"

    return {
        "source": str(input_path),
        "output": str(output_path),
        "line_count": len(lines),
        "mode": mode,
    }


def main() -> int:
    configure_stdio()
    parser = argparse.ArgumentParser(
        description="将长 Markdown 拆分为 AI Native 主入口 + 子文档结构",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 .md 文件或包含 .md 的目录")
    parser.add_argument("-o", "--output-dir", default="ai-native", help="输出目录，默认 ai-native/")
    parser.add_argument("--threshold", type=int, default=500, help="超过多少行视为长文档，默认 500")
    parser.add_argument("--max-lines-per-doc", type=int, default=500, help="拆分后单个子文档目标最大行数，默认 500")
    parser.add_argument("--force", action="store_true", help="即使未超过阈值也生成主入口 + 子文档结构")
    args = parser.parse_args()

    input_path = resolve_path(args.input)
    output_root = resolve_path(args.output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    files, input_root = discover_markdown_files(input_path, output_root)
    if not files:
        print(f"未找到 Markdown 文件：{input_path}", file=sys.stderr)
        return 1

    results = []
    failures = []
    for path in files:
        try:
            result = process_file(path, output_root, input_root, args.threshold, args.max_lines_per_doc, args.force)
            results.append(result)
            print(f"[{result['mode']}] {path} -> {result['output']} ({result['line_count']} 行)")
        except Exception as exc:  # noqa: BLE001
            failures.append({"source": str(path), "error": str(exc)})
            print(f"[failed] {path}: {exc}", file=sys.stderr)

    summary = {
        "success": len(results),
        "failed": len(failures),
        "output_dir": str(output_root),
        "results": results,
        "failures": failures,
    }
    write_text(output_root / "summary.json", json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n完成：成功 {len(results)}，失败 {len(failures)}")
    print(f"汇总文件：{output_root / 'summary.json'}")
    return 0 if not failures else 2


if __name__ == "__main__":
    raise SystemExit(main())
