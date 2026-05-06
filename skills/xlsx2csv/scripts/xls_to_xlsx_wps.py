#!/usr/bin/env python3
from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path


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


def discover_xls_files(raw_inputs: list[str]) -> list[Path]:
    files: list[Path] = []

    if not raw_inputs:
        raw_inputs = ["."]

    for raw in raw_inputs:
        path = resolve_path(raw)

        if not path.exists():
            print(f"警告：路径不存在，已跳过：{path}")
            continue

        if path.is_dir():
            files.extend(
                sorted(
                    (
                        item.resolve()
                        for item in path.rglob("*")
                        if item.is_file()
                        and not item.name.startswith("~$")
                        and item.suffix.lower() == ".xls"
                    ),
                    key=lambda item: str(item.relative_to(path)).casefold(),
                )
            )
            continue

        if path.is_file():
            if path.suffix.lower() == ".xls" and not path.name.startswith("~$"):
                files.append(path)
            else:
                print(f"警告：非 .xls 文件，已跳过：{path}")

    unique: list[Path] = []
    seen: set[str] = set()
    for file_path in files:
        key = str(file_path).casefold()
        if key not in seen:
            seen.add(key)
            unique.append(file_path)
    return unique


def get_output_path(xls_path: Path, output_dir: Path | None) -> Path:
    if output_dir is None:
        return xls_path.with_suffix(".xlsx")
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / xls_path.with_suffix(".xlsx").name


def convert_with_wps_on_windows(xls_path: Path, output_dir: Path | None = None) -> tuple[bool, str]:
    output_path = get_output_path(xls_path, output_dir)
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:
        return False, f"缺少 pywin32 依赖，无法调用 WPS COM：{exc}"

    app = None
    workbook = None
    try:
        pythoncom.CoInitialize()

        for prog_id in ("KET.Application", "ET.Application", "ket.Application", "Excel.Application"):
            try:
                app = win32com.client.DispatchEx(prog_id)
                if app is not None:
                    break
            except Exception:
                continue

        if app is None:
            return False, "无法创建 WPS 表格或 Excel COM 对象。请确认已安装 WPS Spreadsheets 或 Microsoft Excel。"

        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass

        if output_path.exists():
            output_path.unlink()

        workbook = app.Workbooks.Open(str(xls_path))
        try:
            workbook.SaveAs(str(output_path), FileFormat=51)
        except TypeError:
            workbook.SaveAs(str(output_path), 51)

        if not output_path.exists():
            return False, "WPS/Excel 返回成功，但未生成输出文件。"

        return True, f"{xls_path} -> {output_path}"
    except Exception as exc:
        return False, str(exc)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convert_with_libreoffice(xls_path: Path, output_dir: Path | None = None) -> tuple[bool, str]:
    output_path = get_output_path(xls_path, output_dir)
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if not office:
        return False, "非 Windows 环境未找到 soffice/libreoffice，无法转换 .xls。"

    proc = subprocess.run(
        [
            office,
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_path.parent),
            str(xls_path),
        ],
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )

    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        if not err:
            err = "LibreOffice 转换失败。"
        return False, err

    if not output_path.exists():
        return False, "转换命令执行成功但未找到输出文件。"

    return True, f"{xls_path} -> {output_path}"


def main() -> int:
    parser = argparse.ArgumentParser(description="将 .xls 预转换为 .xlsx")
    parser.add_argument("input", nargs="*", help="输入的 .xls 文件或目录")
    parser.add_argument("--output-dir", default=None, help="输出 .xlsx 的目录；默认输出到 .xls 同目录")
    args = parser.parse_args()

    xls_files = discover_xls_files(args.input)
    if not xls_files:
        print("未找到可转换的 .xls 文件。")
        return 0

    is_windows = platform.system().lower().startswith("win")
    output_dir = resolve_path(args.output_dir) if args.output_dir else None
    failures = 0

    for xls_file in xls_files:
        if is_windows:
            ok, msg = convert_with_wps_on_windows(xls_file, output_dir)
        else:
            ok, msg = convert_with_libreoffice(xls_file, output_dir)

        if ok:
            print(f"转换成功：{msg}")
        else:
            failures += 1
            print(f"转换失败：{xls_file}。错误：{msg}", file=sys.stderr)

    return 0 if failures == 0 else 1


if __name__ == "__main__":
    configure_stdio()
    raise SystemExit(main())
