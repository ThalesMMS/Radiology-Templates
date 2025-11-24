#!/usr/bin/env python3
"""
Convert .txt files to Markdown applying formatting rules:
- First non-empty line: bold.
- Last non-empty line: italic.
- Section lines (e.g., exam technique, observed findings, impression) are bold.
"""

import argparse
import sys
from pathlib import Path
from typing import Iterable, List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
TXT_DIR = REPO_ROOT / "Templates_txt"
MD_DIR = REPO_ROOT / "Templates_markdown"

SECTION_PREFIXES = [
    "técnica do exame:",
    "aspectos observados:",
    "impressão:",
    "informe clínico:",
    "indicação clínica:",
    "indicação:",
]


def find_first_last_nonempty(lines: List[str]) -> Optional[tuple[int, int]]:
    """Return indices of first and last non-empty lines, or None if all empty."""
    nonempty_indices = [i for i, line in enumerate(lines) if line.strip()]
    if not nonempty_indices:
        return None
    return nonempty_indices[0], nonempty_indices[-1]


def should_bold_section(line: str) -> bool:
    """Decide if a line should be bolded as a section heading."""
    lowered = line.casefold().strip()
    if any(lowered.startswith(prefix) for prefix in SECTION_PREFIXES):
        return True
    # Generic heuristic: short line ending with ':' looks like a heading.
    if lowered.endswith(":") and len(lowered) <= 120:
        return True
    return False


def format_lines_as_markdown(lines: Iterable[str]) -> List[str]:
    """Apply formatting rules to raw text lines."""
    line_list = list(lines)
    first_last = find_first_last_nonempty(line_list)
    first_idx = last_idx = None
    if first_last is not None:
        first_idx, last_idx = first_last

    output: List[str] = []
    for idx, line in enumerate(line_list):
        stripped = line.strip()
        if not stripped:
            output.append("")
            continue

        is_first = first_idx is not None and idx == first_idx
        is_last = last_idx is not None and idx == last_idx

        text = stripped
        if is_last:
            text = f"*{text}*"
        elif is_first or should_bold_section(stripped):
            text = f"**{text}**"

        output.append(text)

    return output


def convert_txt_file(txt_path: Path, output_dir: Path) -> None:
    """Convert a single .txt file to Markdown."""
    lines = txt_path.read_text(encoding="utf-8").splitlines()
    formatted = format_lines_as_markdown(lines)
    output_dir.mkdir(parents=True, exist_ok=True)
    md_path = output_dir / f"{txt_path.stem}.md"
    md_path.write_text("\n".join(formatted), encoding="utf-8")
    print(f"✓ {txt_path.name} -> {md_path.name}")


def convert_folder(txt_dir: Path, output_dir: Path) -> None:
    """Convert all .txt files in a folder to Markdown."""
    txt_files = sorted(txt_dir.glob("*.txt"))
    if not txt_files:
        print(f"No .txt files found in {txt_dir}")
        return

    for txt_file in txt_files:
        convert_txt_file(txt_file, output_dir)


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Templates_txt/*.txt to Templates_markdown/*.md applying formatting rules."
    )
    parser.add_argument(
        "--txt-dir",
        type=Path,
        default=TXT_DIR,
        help="Source folder for .txt files (default: Templates_txt).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=MD_DIR,
        help="Destination folder for .md files (default: Templates_markdown).",
    )
    return parser.parse_args(argv)


def main(argv=None) -> None:
    args = parse_args(argv)

    if not args.txt_dir.exists():
        raise SystemExit(f"Source folder not found: {args.txt_dir}")

    convert_folder(args.txt_dir, args.output_dir)
    print(f"\n✓ Markdown generated in {args.output_dir}")


if __name__ == "__main__":
    main(sys.argv[1:])
