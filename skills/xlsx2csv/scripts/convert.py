#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlsx2csv 转换脚本

将 .xls/.xlsx 文件转换为一组 CSV 文件：
- 一个索引 CSV（包含各工作表元信息）
- 每个工作表一个 CSV（保留原始网格结构）

支持单文件或目录批量转换。.xls 会先预转换为 .xlsx，再复用同一条 CSV 导出流程。

用法：
    python convert.py <input.xls|input.xlsx or dir> [-o <dir>] [--config <config.yaml>]

依赖：
    pandas, python-calamine, pyyaml
"""

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd
import yaml


def configure_stdio() -> None:
    """脚本日志统一使用 UTF-8，避免中文路径和文件名输出乱码。"""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def resolve_path(raw_path: str) -> Path:
    """解析命令行路径，保留中文等非 ASCII 字符。"""
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    return path.resolve()


def path_key(path: Path) -> str:
    """返回路径比较用的规范化键，兼容大小写不敏感的文件系统。"""
    return os.path.normcase(str(path.resolve())).casefold()


def load_config(config_path: str) -> dict:
    """加载配置文件，返回配置字典。"""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def sanitize_filename(name: str, existing: set = None) -> str:
    """替换文件系统不允许的字符为下划线，如有冲突则追加序号。"""
    safe = re.sub(r'[\\/:*?"<>|]', "_", name)
    if existing is None:
        return safe
    original = safe
    seq = 2
    while safe in existing:
        safe = f"{original}_{seq}"
        seq += 1
    return safe


def get_used_range(df: pd.DataFrame) -> str:
    """
    计算 DataFrame 的有效区域，返回 Excel 风格的范围字符串（如 A1:E10）。
    空 DataFrame 返回空字符串。
    """
    if df.empty or df.shape == (0, 0):
        return ""

    num_rows = df.shape[0]
    num_cols = df.shape[1]

    # 列号转 Excel 列名（A, B, ..., Z, AA, AB, ...）
    def col_to_letter(col_num):
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result

    end_col = col_to_letter(num_cols)
    return f"A1:{end_col}{num_rows}"


def get_sheet_visibility(xlsx_path: str) -> dict:
    """
    通过直接解析 xlsx 的 XML 获取工作表可见性。
    避免 openpyxl 样式解析兼容问题。
    返回 {sheet_name: state} 字典。
    """
    import zipfile
    import xml.etree.ElementTree as ET

    visibility = {}
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        data = zf.read("xl/workbook.xml").decode("utf-8")
        root = ET.fromstring(data)
        ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        for sheet in root.findall(".//ns:sheet", ns):
            name = sheet.get("name")
            state = sheet.get("state", "visible")
            visibility[name] = state

    return visibility


def discover_input_files(input_path: Path) -> list[Path]:
    """
    发现输入的 .xls/.xlsx 文件。

    目录模式递归扫描，跳过 Office 临时文件。如果同目录下同时存在
    `name.xls` 和 `name.xlsx`，优先使用原生 `.xlsx`，避免重复导出。
    """
    allowed_suffixes = {".xls", ".xlsx"}

    if input_path.is_dir():
        entries = sorted(
            (
                file_path.resolve()
                for file_path in input_path.rglob("*")
                if file_path.is_file()
                and not file_path.name.startswith("~$")
                and file_path.suffix.lower() in allowed_suffixes
            ),
            key=lambda file_path: str(file_path.relative_to(input_path)).casefold(),
        )
    elif input_path.is_file():
        if input_path.name.startswith("~$") or input_path.suffix.lower() not in allowed_suffixes:
            return []
        entries = [input_path.resolve()]
    else:
        return []

    xlsx_targets = {
        path_key(file_path.with_suffix(".xlsx"))
        for file_path in entries
        if file_path.suffix.lower() == ".xlsx"
    }

    files = []
    for file_path in entries:
        if file_path.suffix.lower() == ".xls":
            target_xlsx = path_key(file_path.with_suffix(".xlsx"))
            if target_xlsx in xlsx_targets:
                print(
                    f"跳过 .xls，因为同名 .xlsx 已存在：{file_path.name}",
                    file=sys.stderr,
                )
                continue
        files.append(file_path.resolve())
    return files


def convert_xls_via_external_script(xls_path: Path, temp_root: Path) -> tuple[Path | None, str | None]:
    """调用同目录脚本 `xls_to_xlsx_wps.py` 将 `.xls` 预转换为 `.xlsx`。"""
    script_path = Path(__file__).resolve().parent / "xls_to_xlsx_wps.py"

    if not script_path.exists():
        return None, f"预转换脚本不存在：{script_path}"

    temp_root.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(tempfile.mkdtemp(prefix=f"{xls_path.stem}_", dir=temp_root))
    output_path = temp_dir / xls_path.with_suffix(".xlsx").name

    cmd = [sys.executable, str(script_path), str(xls_path), "--output-dir", str(temp_dir)]
    xls = None
    try:
        proc = subprocess.run(
            cmd,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except OSError as exc:
        return None, f"执行预转换脚本失败：{exc}"

    if proc.returncode != 0:
        error = (proc.stderr or proc.stdout or "").strip()
        if not error:
            error = "未知预转换错误"
        return None, error

    if not output_path.exists():
        return None, f"预转换完成但未找到输出 .xlsx：{output_path}"

    return output_path, None


def relative_to_input_root(input_path: Path, input_root: Path | None) -> Path:
    """返回目录输入下的相对路径；单文件输入只返回文件名。"""
    if input_root is None:
        return Path(input_path.name)
    try:
        return input_path.relative_to(input_root)
    except ValueError:
        return Path(input_path.name)


def convert(input_path: str, output_dir: str, config: dict, input_root: Path | None = None) -> None:
    """执行 xls/xlsx → CSV 转换。"""
    source_path = resolve_path(input_path)
    if not source_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{source_path}")

    # 文件名（不含扩展名）作为输出目录名
    relative_source = relative_to_input_root(source_path, input_root)
    stem = source_path.stem
    out_root = resolve_path(output_dir) / relative_source.parent / stem
    out_root.mkdir(parents=True, exist_ok=True)

    encoding = config.get("output", {}).get("encoding", "utf-8-sig")
    include_hidden = config.get("sheet", {}).get("include_hidden_sheets", True)

    workbook_path = source_path
    temp_root = Path.cwd() / "test" / "xlsx2csv_tmp"
    if source_path.suffix.lower() == ".xls":
        print(f"  预转换 .xls 为 .xlsx：{source_path.name}")
        converted_xlsx_path, error = convert_xls_via_external_script(source_path, temp_root)
        if converted_xlsx_path is None:
            raise RuntimeError(f".xls 预转换失败：{error}")
        workbook_path = converted_xlsx_path

    try:
        # 获取工作表可见性
        visibility = get_sheet_visibility(str(workbook_path))

        # 使用 calamine 引擎读取（比 openpyxl 更健壮，不解析样式）
        xls = pd.ExcelFile(str(workbook_path), engine="calamine")
        sheet_names = xls.sheet_names

        index_rows = []
        used_filenames = set()

        for order, name in enumerate(sheet_names, start=1):
            # 根据配置决定是否跳过隐藏 sheet
            state = visibility.get(name, "visible")
            if not include_hidden and state != "visible":
                continue

            # 读取数据：不推断表头，保留所有空白
            df = pd.read_excel(xls, sheet_name=name, header=None, dtype=str)

            # 安全文件名（含冲突检测）
            safe_name = sanitize_filename(name, used_filenames)
            used_filenames.add(safe_name)
            csv_filename = f"{safe_name}.csv"
            csv_path = out_root / csv_filename

            # 导出 CSV
            df.to_csv(
                csv_path,
                index=False,
                header=False,
                encoding=encoding,
                quoting=csv.QUOTE_MINIMAL,
            )

            # 有效区域
            used_range = get_used_range(df)

            # 记录索引信息
            index_rows.append(
                {
                    "工作表顺序": order,
                    "工作表名": name,
                    "导出文件名": csv_filename,
                    "有效区域": used_range,
                }
            )

            status = "hidden" if state != "visible" else "visible"
            print(f"  [{status}] {name} → {csv_filename}  ({used_range or '空'})")

        # 生成索引 CSV
        index_filename = f"{stem}.csv"
        index_path = out_root / index_filename

        index_df = pd.DataFrame(index_rows)
        index_df.to_csv(index_path, index=False, encoding=encoding, quoting=csv.QUOTE_MINIMAL)

        print(f"\n完成！输出目录：{out_root}")
        print(f"  索引文件：{index_filename}")
        print(f"  工作表数：{len(index_rows)}")
    finally:
        if xls is not None:
            try:
                xls.close()
            except Exception:
                pass
        if source_path.suffix.lower() == ".xls" and temp_root.exists():
            shutil.rmtree(temp_root, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="将 xls/xlsx 文件转换为 CSV 集合",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 xls/xlsx 文件路径或包含 xls/xlsx 的目录")
    parser.add_argument(
        "-o", "--output-dir",
        default="csv",
        help="输出目录的父目录（默认为当前工作目录下的 csv/）",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="配置文件路径（默认使用 skill 目录下的 config.yaml）",
    )

    args = parser.parse_args()

    # 确定配置文件路径
    if args.config:
        config_path = args.config
    else:
        # 默认使用脚本所在目录的父目录下的 config.yaml
        script_dir = Path(__file__).resolve().parent
        config_path = script_dir.parent / "config.yaml"

    if not Path(config_path).exists():
        print(f"警告：配置文件不存在：{config_path}，使用默认配置", file=sys.stderr)
        config = {}
    else:
        config = load_config(str(config_path))

    output_dir = args.output_dir

    # 收集待转换文件
    input_path = resolve_path(args.input)
    if input_path.is_dir():
        input_root = input_path
        input_files = discover_input_files(input_path)
        if not input_files:
            print(f"错误：目录中没有找到 .xls/.xlsx 文件：{input_path}", file=sys.stderr)
            sys.exit(1)
    else:
        input_root = None
        input_files = discover_input_files(input_path)
        if not input_files:
            if not input_path.exists():
                print(f"错误：输入路径不存在：{input_path}", file=sys.stderr)
            else:
                print(f"错误：输入不是 .xls/.xlsx 文件：{input_path}", file=sys.stderr)
            sys.exit(1)

    success, failed = 0, 0
    for input_file in input_files:
        print(f"输入文件：{input_file}")
        print(f"输出目录：{output_dir}")
        print()
        try:
            convert(str(input_file), output_dir, config, input_root)
            success += 1
        except Exception as e:
            print(f"错误：转换失败：{input_file} — {e}", file=sys.stderr)
            failed += 1
        if len(input_files) > 1:
            print()

    if len(input_files) > 1:
        print(f"批量完成：成功 {success}，失败 {failed}，共 {len(input_files)} 个文件")


if __name__ == "__main__":
    configure_stdio()
    main()
