# Reports – Template Converter

Automates round-trip conversions between DOCX, Markdown and TXT for medical report templates with consistent formatting rules.

## Features
- `convert_to_docx.py`: builds DOCX files from `Templates_markdown/` with Arial 10, no extra spacing, justified body text, centered first/last lines, last line forced italic size 8.
- `convert_to_markdown.py`: converts every `.docx` in `Templates_docx/` to Markdown, preserving headings, bold, italic and underline; cleans common RTF artifacts when present.
- `convert_to_txt.py`: converts Markdown to TXT (strips `*` and `#`), or use `--from-docx` to convert DOCX → temporary Markdown → TXT.
- `convert_txt_to_markdown.py`: converts TXT back to Markdown applying rules (first line bold, last line italic, section headers like exam technique bold).

## Requirements
- Python 3.8+
- macOS or Linux shell (tested on macOS 15)
- `python-docx`, `striprtf` (auto-installed on demand by the markdown converter)

## Usage
- DOCX → Markdown:
  ```bash
  python convert_to_markdown.py
  ```
  Outputs to `Templates_markdown/`.

- Markdown → DOCX:
  ```bash
  python convert_to_docx.py
  ```
  Outputs to `Templates_docx/` (created if missing).

- Markdown → TXT (default) or DOCX → TXT:
  ```bash
  # from existing markdown
  python convert_to_txt.py

  # convert docx to markdown in a temp folder, then to txt
  python convert_to_txt.py --from-docx
  ```
  Outputs to `Templates_txt/`.

- TXT → Markdown (reapplies basic formatting):
  ```bash
  python convert_txt_to_markdown.py
  ```
  Outputs to `Templates_markdown/`.

Scripts print progress and overwrite existing outputs; commit or back up generated files as needed.

## Repository Structure
- `convert_to_docx.py` – Markdown → DOCX generator with alignment/font rules.
- `convert_to_markdown.py` – DOCX/RTF → Markdown converter.
- `convert_to_txt.py` – Markdown (or DOCX via temp markdown) → TXT converter.
- `convert_txt_to_markdown.py` – TXT → Markdown converter with heading/first/last-line rules.
- `Templates_markdown/` – source Markdown templates.
- `Templates_docx/` – DOCX output from Markdown (and DOCX input for md conversion).
- `Templates_txt/` – TXT output.

## Notes
- Markdown parsing in `convert_to_docx.py` supports bold/italic markers and centers the first and last non-empty lines.
- DOCX → Markdown relies on Word styles (e.g., `Heading 1`) to infer heading levels; keep templates consistent.
- RTF parsing uses heuristics; review new outputs when adding unfamiliar RTFs.
