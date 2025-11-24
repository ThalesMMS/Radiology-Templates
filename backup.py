#!/usr/bin/env python3
"""
Move any files in Templates_docx, Templates_markdown, or Templates_txt that are
not listed in reports_index.json into backup/, preserving folder structure.
"""

import json
import shutil
from pathlib import Path
from typing import Dict, Iterable, Set

BASE_DIR = Path(__file__).parent
TARGETS: Dict[str, str] = {
    "Templates_docx": "*.docx",
    "Templates_markdown": "*.md",
    "Templates_txt": "*.txt",
}
INDEX_PATH = BASE_DIR / "reports_index.json"
BACKUP_DIR = BASE_DIR / "backup"


def load_index() -> Dict[str, Set[str]]:
    if not INDEX_PATH.exists():
        raise SystemExit("Index file not found. Run generate_index.py first.")

    data = json.loads(INDEX_PATH.read_text(encoding="utf-8"))
    indexed: Dict[str, Set[str]] = {}
    for folder_name in TARGETS.keys():
        indexed[folder_name] = set(data.get(folder_name, []))
    return indexed


def iter_files(folder: Path, pattern: str) -> Iterable[Path]:
    return (p for p in folder.glob(pattern) if p.is_file())


def move_unindexed(indexed: Dict[str, Set[str]]) -> None:
    BACKUP_DIR.mkdir(exist_ok=True)
    moved = 0
    for folder_name, pattern in TARGETS.items():
        folder = BASE_DIR / folder_name
        if not folder.exists():
            print(f"Skipping missing folder: {folder}")
            continue

        expected = indexed.get(folder_name, set())
        for path in iter_files(folder, pattern):
            rel_path = str(path.relative_to(BASE_DIR))
            if rel_path in expected:
                continue

            dest = BACKUP_DIR / rel_path
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                print(f"Skip {rel_path}: destination already exists in backup.")
                continue

            shutil.move(str(path), str(dest))
            moved += 1
            print(f"Moved {rel_path} -> {dest.relative_to(BASE_DIR)}")

    print(f"\nDone. Files moved: {moved}")


def main() -> None:
    indexed = load_index()
    move_unindexed(indexed)


if __name__ == "__main__":
    main()
