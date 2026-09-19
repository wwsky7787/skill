#!/usr/bin/env python3
"""
Everything 命令行搜索封装脚本。

自动查找 es.exe，构造命令行参数，执行搜索并格式化输出。
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Tuple


COMMON_ES_PATHS = [
    r"C:\Program Files\Everything\es.exe",
    r"C:\Program Files (x86)\Everything\es.exe",
]


def find_es() -> Optional[Path]:
    """查找 es.exe 可执行文件。"""
    # 1. PATH
    es_in_path = shutil.which("es.exe")
    if es_in_path:
        return Path(es_in_path)

    # 2. 常见安装目录
    for p in COMMON_ES_PATHS:
        path = Path(p)
        if path.exists():
            return path

    # 3. skill 自带的 bin 目录
    script_dir = Path(__file__).resolve().parent
    bundled_bin = script_dir.parent / "bin" / "es.exe"
    if bundled_bin.exists():
        return bundled_bin.resolve()

    # 4. 当前目录及上级目录（可能用户把 es.exe 放在 skill 目录）
    for base in [script_dir, Path.cwd()]:
        candidate = base / "es.exe"
        if candidate.exists():
            return candidate.resolve()

    return None


def build_command(args: argparse.Namespace, es_path: Path) -> List[str]:
    """根据参数构造 es.exe 命令行。"""
    cmd = [str(es_path)]

    # 查询文本
    if args.query:
        cmd.append(args.query)

    if args.regex:
        cmd.append("-regex")
    if args.case:
        cmd.append("-case")
    if args.whole_word:
        cmd.append("-whole-word")
    if args.match_path:
        cmd.append("-match-path")

    if args.max_results is not None:
        cmd.extend(["-n", str(args.max_results)])

    if args.offset is not None:
        cmd.extend(["-offset", str(args.offset)])

    if args.path:
        cmd.extend(["-path", args.path])

    if args.parent:
        cmd.extend(["-parent", args.parent])

    if args.sort:
        cmd.extend(["-sort", args.sort])

    if args.sort_ascending:
        cmd.append("-sort-ascending")
    elif args.sort_descending:
        cmd.append("-sort-descending")

    # 分栏显示
    columns = ["name"] if args.format != "list" else []
    display_columns = {
        "size": args.show_size,
        "date-modified": args.show_date_modified,
        "date-created": args.show_date_created,
        "date-accessed": args.show_date_accessed,
        "attributes": args.show_attributes,
        "path": args.show_path,
    }
    for col, enabled in display_columns.items():
        if enabled:
            columns.append(col)

    for col in columns:
        cmd.append(f"-{col}")

    # 导出到文件（优先级高于屏幕输出）
    if args.output:
        if args.format == "csv":
            cmd.extend(["-export-csv", args.output])
        elif args.format == "efu":
            cmd.extend(["-export-efu", args.output])
        elif args.format == "m3u":
            cmd.extend(["-export-m3u", args.output])
        elif args.format == "m3u8":
            cmd.extend(["-export-m3u8", args.output])
        else:
            cmd.extend(["-export-txt", args.output])

    return cmd


def run_search(cmd: List[str]) -> Tuple[int, str, str]:
    """运行 es.exe 命令并返回结果。"""
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        return result.returncode, result.stdout, result.stderr
    except FileNotFoundError as e:
        return 8, "", f"找不到 es.exe: {e}"
    except Exception as e:
        return 8, "", f"执行 es.exe 时出错: {e}"


def format_results(stdout: str, output_format: str, returncode: int) -> dict:
    """格式化搜索结果。"""
    lines = [line for line in stdout.splitlines() if line.strip()]

    if output_format == "json":
        return {
            "success": returncode == 0,
            "count": len(lines),
            "results": lines,
        }

    if output_format == "list":
        return {
            "success": returncode == 0,
            "count": len(lines),
            "results": lines,
            "text": "\n".join(lines),
        }

    # text / csv / efu / m3u / m3u8 都作为文本处理
    return {
        "success": returncode == 0,
        "count": len(lines),
        "text": stdout,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="通过 Everything 命令行接口搜索本地文件。"
    )
    parser.add_argument("query", nargs="?", default="", help="搜索查询（Everything 查询语法）")
    parser.add_argument("-n", "--max-results", type=int, help="最大结果数")
    parser.add_argument("--offset", type=int, help="结果偏移量")
    parser.add_argument("-r", "--regex", action="store_true", help="使用正则表达式")
    parser.add_argument("-i", "--case", action="store_true", help="区分大小写")
    parser.add_argument("-w", "--whole-word", action="store_true", help="全字匹配")
    parser.add_argument("-p", "--match-path", action="store_true", help="匹配完整路径")
    parser.add_argument("--path", help="限制搜索路径")
    parser.add_argument("--parent", help="搜索指定父目录下的子文件")
    parser.add_argument(
        "--sort",
        choices=[
            "name", "path", "size", "extension",
            "date-created", "date-modified", "date-accessed",
            "attributes", "file-list-file-name", "run-count",
            "date-recently-changed", "date-run",
        ],
        help="排序字段",
    )
    parser.add_argument("--sort-ascending", action="store_true", help="升序")
    parser.add_argument("--sort-descending", action="store_true", help="降序")

    parser.add_argument("--show-size", action="store_true", help="显示大小")
    parser.add_argument("--show-date-modified", action="store_true", help="显示修改日期")
    parser.add_argument("--show-date-created", action="store_true", help="显示创建日期")
    parser.add_argument("--show-date-accessed", action="store_true", help="显示访问日期")
    parser.add_argument("--show-attributes", action="store_true", help="显示属性")
    parser.add_argument("--show-path", action="store_true", help="显示完整路径")

    parser.add_argument(
        "-f",
        "--format",
        choices=["text", "json", "list", "csv", "efu", "m3u", "m3u8"],
        default="text",
        help="输出格式",
    )
    parser.add_argument("-o", "--output", help="导出到文件路径")

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    es_path = find_es()
    if not es_path:
        error_msg = (
            "错误：未找到 es.exe。\n"
            "请从 https://www.voidtools.com/zh-cn/downloads/ 下载 Command-line Interface (ES)，\n"
            "并将其放入 Everything 安装目录或系统 PATH 中。"
        )
        if args.format == "json":
            print(json.dumps({"success": False, "error": error_msg}, ensure_ascii=False))
        else:
            print(error_msg, file=sys.stderr)
        return 8

    cmd = build_command(args, es_path)
    returncode, stdout, stderr = run_search(cmd)

    if returncode != 0:
        error_msg = stderr.strip() or f"es.exe 返回错误码 {returncode}"
        if args.format == "json":
            print(json.dumps({"success": False, "error": error_msg}, ensure_ascii=False))
        else:
            print(error_msg, file=sys.stderr)
        return returncode

    result = format_results(stdout, args.format, returncode)

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.format == "list":
        print("\n".join(result["results"]))
    else:
        if args.output:
            print(f"搜索结果已导出到: {args.output}")
            print(f"共 {result['count']} 条结果")
        else:
            print(result["text"])

    return 0


if __name__ == "__main__":
    sys.exit(main())
