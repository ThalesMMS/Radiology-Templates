#!/usr/bin/env python3
"""
Build an index of laudo files in Templates_docx, Templates_markdown, and Templates_txt.
The index is written to reports_index.json in the repo root.
"""

import json
from pathlib import Path
from typing import Dict, List

BASE_DIR = Path(__file__).parent
TARGETS = {
    "Templates_docx": "*.docx",
    "Templates_markdown": "*.md",
    "Templates_txt": "*.txt",
}
INDEX_PATH = BASE_DIR / "reports_index.json"


def collect_index() -> Dict[str, List[str]]:
    """Collect file paths per folder, relative to repo root."""
    index: Dict[str, List[str]] = {}
    for folder_name, pattern in TARGETS.items():
        folder = BASE_DIR / folder_name
        if not folder.exists():
            print(f"Skipping missing folder: {folder}")
            index[folder_name] = []
            continue

        files = sorted(
            str(path.relative_to(BASE_DIR))
            for path in folder.glob(pattern)
            if path.is_file()
        )
        index[folder_name] = files
        print(f"{folder_name}: {len(files)} files")

    return index


def main() -> None:
    index = collect_index()
    INDEX_PATH.write_text(
        json.dumps(index, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"\nIndex written to {INDEX_PATH}")


if __name__ == "__main__":
    main()
