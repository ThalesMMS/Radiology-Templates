#!/usr/bin/env python3
"""
Generate .txt files from markdown (default flow) or, optionally, from .docx by
first converting them to markdown in a temporary folder.
"""

import argparse
import sys
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).parent
MD_DIR = BASE_DIR / "Templates_markdown"
DOCX_DIR = BASE_DIR / "Templates_docx"
TXT_DIR = BASE_DIR / "Templates_txt"


def clean_markdown_text(text: str) -> str:
    """Remove basic markdown markers (* and #) while keeping line breaks."""
    return text.replace("*", "").replace("#", "")


def convert_md_file(md_path: Path, output_dir: Path) -> None:
    """Convert a markdown file to txt, stripping markers."""
    output_dir.mkdir(parents=True, exist_ok=True)
    txt_path = output_dir / f"{md_path.stem}.txt"
    content = md_path.read_text(encoding="utf-8")
    cleaned = clean_markdown_text(content)
    txt_path.write_text(cleaned, encoding="utf-8")


def convert_markdown_folder(md_dir: Path, output_dir: Path) -> None:
    """Convert all .md files in a folder."""
    md_files = sorted(md_dir.glob("*.md"))
    if not md_files:
        print(f"No .md files found in {md_dir}")
        return

    for md_file in md_files:
        convert_md_file(md_file, output_dir)
        print(f"✓ {md_file.name} -> {md_file.stem}.txt")


def convert_from_docx(output_dir: Path) -> None:
    """Alternate flow: convert .docx to temporary markdown, then to txt."""
    try:
        from convert_to_markdown import convert_docx_to_markdown
    except Exception as exc:  # pragma: no cover - defensive import
        print("Erro ao importar convert_docx_to_markdown de convert_to_markdown.py")
        raise SystemExit(exc)

    docx_files = sorted(DOCX_DIR.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {DOCX_DIR}")
        return

    with tempfile.TemporaryDirectory(prefix="md_tmp_") as tmp_dir:
        tmp_md_dir = Path(tmp_dir)
        for docx_file in docx_files:
            markdown_content = convert_docx_to_markdown(docx_file)
            md_output = tmp_md_dir / f"{docx_file.stem}.md"
            md_output.write_text(markdown_content, encoding="utf-8")
            print(f"Generated temporary: {md_output.name}")

        convert_markdown_folder(tmp_md_dir, output_dir)
        # TemporaryDirectory se encarrega de apagar tmp_md_dir


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Converts markdown -> txt (default). "
            "Use --from-docx to convert docx -> markdown -> txt."
        )
    )
    parser.add_argument(
        "--from-docx",
        action="store_true",
        help="Generate txt from .docx (with an intermediate markdown conversion).",
    )
    return parser.parse_args(argv)


def main(argv=None) -> None:
    args = parse_args(argv)

    if args.from_docx:
        convert_from_docx(TXT_DIR)
    else:
        if not MD_DIR.exists():
            raise SystemExit(f"Source folder not found: {MD_DIR}")
        convert_markdown_folder(MD_DIR, TXT_DIR)

    print(f"\n✓ Arquivos gerados em {TXT_DIR}")


if __name__ == "__main__":
    main(sys.argv[1:])
