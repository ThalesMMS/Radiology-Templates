#!/usr/bin/env python3
"""
test_equivalence.py

Compares outputs from the original Python scripts with the Rust binaries
to ensure they match whenever possible.

Suggested use:

    python test_equivalence.py \
        --project-root . \
        --rust-bin-dir ./rust_converters/target/debug

Requires:
  - Python 3
  - Standard modules (subprocess, filecmp, tempfile, shutil, etc.)
"""

import argparse
import filecmp
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PY_SRC_DIR = "python_src"
PY_SCRIPTS = [
    "convert_to_docx.py",
    "convert_to_markdown.py",
    "convert_to_txt.py",
    "convert_txt_to_markdown.py",
]

RUST_BINS = [
    "convert_to_docx",
    "convert_to_markdown",
    "convert_to_txt",
    "convert_txt_to_markdown",
]


def run_cmd(cmd, cwd):
    """Run a command and raise if it exits with a non-zero code."""
    print(f"[CMD] ({cwd})", " ".join(cmd))
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if result.returncode != 0:
        print("STDOUT:\n", result.stdout)
        print("STDERR:\n", result.stderr, file=sys.stderr)
        raise RuntimeError(f"Command failed: {' '.join(cmd)}")
    return result


def copy_tree(src: Path, dst: Path):
    """Copy a whole folder (only if it exists)."""
    if not src.exists():
        return
    for item in src.iterdir():
        target = dst / item.name
        if item.is_dir():
            shutil.copytree(item, target)
        else:
            shutil.copy2(item, target)


def collect_files(dir_path: Path, ext: str):
    """Return a sorted list of files with the given extension in a folder (non-recursive)."""
    if not dir_path.exists():
        return []
    return sorted(
        p for p in dir_path.iterdir()
        if p.is_file() and p.suffix.lower() == ext.lower()
    )


def compare_text_dirs(dir_a: Path, dir_b: Path, ext: str) -> bool:
    """Compare all files with a given extension in two folders (name + content)."""
    files_a = collect_files(dir_a, ext)
    files_b = collect_files(dir_b, ext)

    names_a = [p.name for p in files_a]
    names_b = [p.name for p in files_b]

    if names_a != names_b:
        print(f"FAIL: file names {ext} differ:")
        print("Python:", names_a)
        print("Rust  :", names_b)
        return False

    ok = True
    for fa, fb in zip(files_a, files_b):
        if not filecmp.cmp(fa, fb, shallow=False):
            print(f"FAIL: content differs in {fa.name}")
            ok = False
        else:
            print(f"OK: {fa.name}")
    return ok


def rust_bin_paths(rust_bin_dir: Path):
    """Return the expected binary paths (considering .exe on Windows)."""
    suffix = ".exe" if os.name == "nt" else ""
    return [rust_bin_dir / f"{name}{suffix}" for name in RUST_BINS]


def ensure_rust_binaries(project_root: Path, rust_bin_dir: Path):
    """Ensure Rust binaries exist; runs `cargo build` if anything is missing."""
    missing_before = [p for p in rust_bin_paths(rust_bin_dir) if not p.exists()]
    if not missing_before:
        return

    cargo_dir = project_root / "rust_converters"
    if not cargo_dir.exists():
        raise FileNotFoundError(f"Could not find the Rust project folder at {cargo_dir}")

    print("Missing Rust binaries; running `cargo build`...")
    run_cmd(["cargo", "build"], cwd=cargo_dir)

    missing_after = [p for p in rust_bin_paths(rust_bin_dir) if not p.exists()]
    if missing_after:
        missing_names = ", ".join(p.name for p in missing_after)
        raise FileNotFoundError(
            f"Binaries still missing after build: {missing_names}"
        )


def copy_python_sources(project_root: Path, target_dir: Path):
    """Copy python_src package and entrypoint scripts to a temp dir."""
    shutil.copytree(project_root / PY_SRC_DIR, target_dir / PY_SRC_DIR)
    for script in PY_SCRIPTS:
        shutil.copy2(project_root / script, target_dir / script)


def setup_temp_env(project_root: Path, py_or_rs: str, rust_bin_dir: Path) -> Path:
    """
    Create a temporary directory containing:
      - Python scripts or Rust binaries (depending on py_or_rs).
      - Templates_* subfolders copied from the project.
    Returns the temp dir path.
    """
    tmp = Path(tempfile.mkdtemp(prefix=f"equiv_{py_or_rs}_"))
    print(f"Temp {py_or_rs} env:", tmp)

    # Copy data folders
    for folder in ("Templates_markdown", "Templates_docx", "Templates_txt"):
        src = project_root / folder
        dst = tmp / folder
        if src.exists():
            shutil.copytree(src, dst)

    if py_or_rs == "py":
        # Copy Python sources and entrypoints
        copy_python_sources(project_root, tmp)
    else:
        # Copy Python sources (needed for convert_to_markdown.py in this flow)
        copy_python_sources(project_root, tmp)
        # Copy Rust binaries (only the ones that exist)
        for bin_name in RUST_BINS:
            src_bin = rust_bin_dir / bin_name
            if os.name == "nt":
                src_bin = rust_bin_dir / (bin_name + ".exe")
            if src_bin.exists():
                shutil.copy2(src_bin, tmp / src_bin.name)

    return tmp


def test_convert_to_markdown(project_root: Path, rust_bin_dir: Path) -> bool:
    print("\n=== Test: convert_to_markdown (.docx/.rtf -> .md) ===")
    tmp_py = setup_temp_env(project_root, "py", rust_bin_dir)
    tmp_rs = setup_temp_env(project_root, "rs", rust_bin_dir)

    # Python
    run_cmd([sys.executable, "convert_to_markdown.py"], cwd=tmp_py)

    # Rust
    rs_bin = tmp_rs / ("convert_to_markdown.exe" if os.name == "nt" else "convert_to_markdown")
    run_cmd([str(rs_bin)], cwd=tmp_rs)

    py_md_dir = tmp_py / "Templates_markdown"
    rs_md_dir = tmp_rs / "Templates_markdown"

    return compare_text_dirs(py_md_dir, rs_md_dir, ".md")


def test_convert_to_txt_markdown_flow(project_root: Path, rust_bin_dir: Path) -> bool:
    print("\n=== Test: convert_to_txt (markdown -> txt) ===")
    tmp_py = setup_temp_env(project_root, "py", rust_bin_dir)
    tmp_rs = setup_temp_env(project_root, "rs", rust_bin_dir)

    # Python
    run_cmd([sys.executable, "convert_to_txt.py"], cwd=tmp_py)

    # Rust
    rs_bin = tmp_rs / ("convert_to_txt.exe" if os.name == "nt" else "convert_to_txt")
    run_cmd([str(rs_bin)], cwd=tmp_rs)

    py_txt_dir = tmp_py / "Templates_txt"
    rs_txt_dir = tmp_rs / "Templates_txt"

    return compare_text_dirs(py_txt_dir, rs_txt_dir, ".txt")


def test_convert_to_txt_from_docx(project_root: Path, rust_bin_dir: Path) -> bool:
    print("\n=== Test: convert_to_txt --from-docx (docx -> md -> txt) ===")
    tmp_py = setup_temp_env(project_root, "py", rust_bin_dir)
    tmp_rs = setup_temp_env(project_root, "rs", rust_bin_dir)

    # Python
    run_cmd([sys.executable, "convert_to_txt.py", "--from-docx"], cwd=tmp_py)

    # Rust
    rs_bin = tmp_rs / ("convert_to_txt.exe" if os.name == "nt" else "convert_to_txt")
    run_cmd([str(rs_bin), "--from-docx"], cwd=tmp_rs)

    py_txt_dir = tmp_py / "Templates_txt"
    rs_txt_dir = tmp_rs / "Templates_txt"

    return compare_text_dirs(py_txt_dir, rs_txt_dir, ".txt")


def test_convert_txt_to_markdown(project_root: Path, rust_bin_dir: Path) -> bool:
    print("\n=== Test: convert_txt_to_markdown (txt -> md) ===")
    tmp_py = setup_temp_env(project_root, "py", rust_bin_dir)
    tmp_rs = setup_temp_env(project_root, "rs", rust_bin_dir)

    # Python
    run_cmd([sys.executable, "convert_txt_to_markdown.py"], cwd=tmp_py)

    # Rust
    rs_bin = tmp_rs / ("convert_txt_to_markdown.exe" if os.name == "nt" else "convert_txt_to_markdown")
    run_cmd([str(rs_bin)], cwd=tmp_rs)

    py_md_dir = tmp_py / "Templates_markdown"
    rs_md_dir = tmp_rs / "Templates_markdown"

    return compare_text_dirs(py_md_dir, rs_md_dir, ".md")


def test_convert_to_docx_roundtrip(project_root: Path, rust_bin_dir: Path) -> bool:
    """
    Indirect equivalence test for convert_to_docx:

      Templates_markdown/*.md
        -> DOCX (Python or Rust)
        -> convert_to_markdown.py (Python)
        -> Templates_markdown/*.md (new)

    We compare these final Markdown files between the two flows.
    """
    print("\n=== Test: convert_to_docx (roundtrip via convert_to_markdown.py) ===")

    # Python environment
    tmp_py = setup_temp_env(project_root, "py", rust_bin_dir)
    # In this env we need both convert_to_docx.py and convert_to_markdown.py
    run_cmd([sys.executable, "convert_to_docx.py"], cwd=tmp_py)
    # Then convert the generated DOCX back to md
    run_cmd([sys.executable, "convert_to_markdown.py"], cwd=tmp_py)
    py_md_dir = tmp_py / "Templates_markdown"

    # Rust environment
    tmp_rs = setup_temp_env(project_root, "rs", rust_bin_dir)
    rs_bin = tmp_rs / ("convert_to_docx.exe" if os.name == "nt" else "convert_to_docx")
    run_cmd([str(rs_bin)], cwd=tmp_rs)
    # Now always use the Python convert_to_markdown.py script
    run_cmd([sys.executable, "convert_to_markdown.py"], cwd=tmp_rs)
    rs_md_dir = tmp_rs / "Templates_markdown"

    return compare_text_dirs(py_md_dir, rs_md_dir, ".md")


def main():
    parser = argparse.ArgumentParser(
        description="Compares outputs from the Python scripts and the Rust binaries."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=REPO_ROOT,
        help="Project root containing Python scripts and Templates_* folders.",
    )
    parser.add_argument(
        "--rust-bin-dir",
        type=Path,
        default=Path("rust_converters/target/debug"),
        help="Folder where the compiled Rust binaries live.",
    )

    args = parser.parse_args()
    project_root: Path = args.project_root.resolve()
    rust_bin_dir: Path = args.rust_bin_dir.resolve()

    print("Project:", project_root)
    print("Rust binaries:", rust_bin_dir)

    ensure_rust_binaries(project_root, rust_bin_dir)

    tests = [
        ("convert_to_markdown", test_convert_to_markdown),
        ("convert_to_txt (markdown)", test_convert_to_txt_markdown_flow),
        ("convert_to_txt (--from-docx)", test_convert_to_txt_from_docx),
        ("convert_txt_to_markdown", test_convert_txt_to_markdown),
        ("convert_to_docx (roundtrip)", test_convert_to_docx_roundtrip),
    ]

    all_ok = True
    for name, func in tests:
        try:
            ok = func(project_root, rust_bin_dir)
        except Exception as exc:
            print(f"[{name}] ERROR: {exc}")
            ok = False
        if not ok:
            all_ok = False
            print(f"[{name}] -> FAIL")
        else:
            print(f"[{name}] -> OK")

    if not all_ok:
        sys.exit(1)
    print("\nAll tests passed.")
    sys.exit(0)


if __name__ == "__main__":
    main()
