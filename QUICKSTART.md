# Quickstart

## Prerequisites
- Python 3.8+
- Rust toolchain (for the Rust binaries)
- macOS or Linux shell

## Install Python deps
Most scripts auto-install `python-docx` and `striprtf` on first run. If you want to preinstall:
```bash
python -m pip install python-docx striprtf
```

## Build Rust binaries
```bash
cd rust_converters
cargo build
```
Binaries are placed in `rust_converters/target/debug`.

## Run conversions (Python)
Use the unified entrypoint:
```bash
# DOCX -> Markdown
python run.py convert_to_markdown

# Markdown -> DOCX
python run.py convert_to_docx

# Markdown -> TXT (or DOCX -> TXT via temp markdown)
python run.py convert_to_txt [--from-docx]

# TXT -> Markdown
python run.py convert_txt_to_markdown

# Update index
python run.py generate_index

# Move unindexed files to backup/
python run.py backup
```

## Rust equivalents
After building, you can run the Rust binaries directly from `rust_converters/target/debug/`:
```bash
./convert_to_markdown
./convert_to_docx
./convert_to_txt [--from-docx]
./convert_txt_to_markdown
./generate_index
./backup
```

## Validate parity
From the repo root:
```bash
python test_equivalence.py
```
This compares Python outputs to Rust outputs (including `generate_index` and `backup`) and should report all tests as passed.
