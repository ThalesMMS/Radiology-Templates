# Reports – Template Converter

Automates conversion of DOCX/RTF medical report templates into consistently formatted Markdown.

## Features
- Converts every `.docx` and `.rtf` file inside `Templates/` to Markdown while preserving headings, bold, italic and underlined text.
- Cleans common RTF artifacts and automatically formats detected report sections.
- Saves the generated Markdown files under `Templates_markdown/`, keeping the original directory untouched.
- Installs the required Python packages (`python-docx`, `striprtf`) on demand the first time the script runs.

## Requirements
- Python 3.8+
- macOS or Linux shell (tested on macOS 15)
- `python-docx` and `striprtf` (auto-installed if missing)

## Usage
1. Place any DOCX or RTF templates inside `Templates/`.
2. Run the converter from the repo root:
   ```bash
   python convert_to_markdown.py
   ```
3. Check the generated Markdown files under `Templates_markdown/`.

The script prints progress for each file and reports the output path on success. Existing Markdown files are overwritten, so commit or back them up if needed.

## Repository Structure
- `convert_to_markdown.py` – main converter script.
- `Templates/` – source templates (treated as vendored for GitHub language stats).
- `Templates_markdown/` – auto-generated Markdown output (also vendored).

## Notes
- The script relies on style names (e.g., `Heading 1`) to infer Markdown heading levels. Ensure templates use consistent Word styles for best results.
- RTF parsing uses heuristics; review the generated Markdown when adding new RTF sources with unusual formatting.

