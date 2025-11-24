#!/usr/bin/env python3
"""
Unified CLI entry point for the Python tools.

Usage:
  python run.py convert_to_docx
  python run.py convert_to_markdown
  python run.py convert_to_txt [--from-docx]
  python run.py convert_txt_to_markdown [--txt-dir DIR] [--output-dir DIR]
  python run.py generate_index
  python run.py backup
"""

import argparse
import sys
from typing import List


def dispatch(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description="Run report converter tools.")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("convert_to_docx", help="Markdown -> DOCX")
    sub.add_parser("convert_to_markdown", help="DOCX/RTF -> Markdown")

    to_txt = sub.add_parser("convert_to_txt", help="Markdown (or DOCX via temp) -> TXT")
    to_txt.add_argument(
        "--from-docx",
        action="store_true",
        help="Convert DOCX -> Markdown -> TXT",
    )

    txt_to_md = sub.add_parser("convert_txt_to_markdown", help="TXT -> Markdown")
    txt_to_md.add_argument("--txt-dir", type=str, default=None, help="TXT source dir")
    txt_to_md.add_argument("--output-dir", type=str, default=None, help="Markdown output dir")

    sub.add_parser("generate_index", help="Build reports_index.json")
    sub.add_parser("backup", help="Move unindexed files to backup/")

    args, rest = parser.parse_known_args(argv)

    if args.command == "convert_to_docx":
        from python_src.convert_to_docx import main as cmd
        cmd()
        return 0

    if args.command == "convert_to_markdown":
        from python_src.convert_to_markdown import main as cmd
        cmd()
        return 0

    if args.command == "convert_to_txt":
        from python_src.convert_to_txt import main as cmd
        cmd(rest)
        return 0

    if args.command == "convert_txt_to_markdown":
        from python_src.convert_txt_to_markdown import main as cmd
        cmd(rest)
        return 0

    if args.command == "generate_index":
        from python_src.generate_index import main as cmd
        cmd()
        return 0

    if args.command == "backup":
        from python_src.backup import main as cmd
        cmd()
        return 0

    parser.error(f"Unknown command {args.command}")
    return 1


if __name__ == "__main__":
    sys.exit(dispatch(sys.argv[1:]))
