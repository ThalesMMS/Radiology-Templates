# Reports – Template Converter

Automates round-trip conversions between DOCX, Markdown and TXT for medical report templates with consistent formatting rules.

## Features
- Ships with radiology report templates plus scripts (Python and Rust) to generate DOCX/Markdown/TXT variants.
- `convert_to_docx`: builds DOCX files from `Templates_markdown/` with Arial 10, no extra spacing, justified body text, centered first/last lines, last line forced italic size 8.
- `convert_to_markdown`: converts every `.docx` in `Templates_docx/` to Markdown, preserving headings, bold, italic and underline; cleans common RTF artifacts when present.
- `convert_to_txt`: converts Markdown to TXT (strips `*` and `#`), or use `--from-docx` to convert DOCX → temporary Markdown → TXT.
- `convert_txt_to_markdown`: converts TXT back to Markdown applying rules (first line bold, last line italic, section headers like exam technique bold).
- `generate_index`: builds `reports_index.json` listing files in `Templates_docx`, `Templates_markdown`, and `Templates_txt`.
- `backup`: moves any files not present in `reports_index.json` from those folders into `backup/`, preserving structure.

## Requirements
- Python 3.8+
- macOS or Linux shell (tested on macOS 15)
- `python-docx`, `striprtf` (auto-installed on demand by the markdown converter)

## Usage (Python via unified entrypoint `run.py`)
- DOCX → Markdown:
  ```bash
  python run.py convert_to_markdown
  ```
  Outputs to `Templates_markdown/`.

- Markdown → DOCX:
  ```bash
  python run.py convert_to_docx
  ```
  Outputs to `Templates_docx/` (created if missing).

- Markdown → TXT (default) or DOCX → TXT:
  ```bash
  # from existing markdown
  python run.py convert_to_txt

  # convert docx to markdown in a temp folder, then to txt
  python run.py convert_to_txt --from-docx
  ```
  Outputs to `Templates_txt/`.

- TXT → Markdown (reapplies basic formatting):
  ```bash
  python run.py convert_txt_to_markdown
  ```
  Outputs to `Templates_markdown/`.

- Build/update index of reports:
  ```bash
  python run.py generate_index
  ```
  Outputs to `reports_index.json`.

- Move unindexed files to backup:
  ```bash
  python run.py backup
  ```
  Moves to `backup/` (mirrors source subfolders).

Scripts print progress and overwrite existing outputs; commit or back up generated files as needed.

Tip: run `python generate_index.py` before `python backup.py` to ensure the backup check uses a fresh index.

## Repository Structure
- `python_src/` – Python sources; invoke via `python run.py <command>`.
  - `convert_to_docx.py` – Markdown → DOCX generator with alignment/font rules.
  - `convert_to_markdown.py` – DOCX/RTF → Markdown converter.
  - `convert_to_txt.py` – Markdown (or DOCX via temp markdown) → TXT converter.
  - `convert_txt_to_markdown.py` – TXT → Markdown converter with heading/first/last-line rules.
  - `generate_index.py` – creates `reports_index.json` for DOCX/Markdown/TXT folders.
  - `backup.py` – moves files not present in `reports_index.json` into `backup/`.
- `run.py` – unified Python CLI entrypoint to all tools.
- `rust_converters/` – Rust implementation of all tools (converters, `generate_index`, `backup`), binaries land in `rust_converters/target/debug`.
- `Templates_markdown/` – source Markdown templates.
- `Templates_docx/` – DOCX output from Markdown (and DOCX input for md conversion).
- `Templates_txt/` – TXT output.

## Notes
- Markdown parsing in `convert_to_docx.py` supports bold/italic markers and centers the first and last non-empty lines.
- DOCX → Markdown relies on Word styles (e.g., `Heading 1`) to infer heading levels; keep templates consistent.
- RTF parsing uses heuristics; review new outputs when adding unfamiliar RTFs.
- Rust parity: `cargo build` inside `rust_converters/` produces binaries equivalent to the Python scripts; `python test_equivalence.py` checks output identity across all tools (including `generate_index` and `backup`).
