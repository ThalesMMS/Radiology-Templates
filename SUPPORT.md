# Support

## Before opening an issue

Please check the README and [QUICKSTART.md](QUICKSTART.md) first.

When reporting a problem, include:

- which workflow failed (`convert_to_docx`, `convert_to_markdown`, `convert_to_txt`, `convert_txt_to_markdown`, `generate_index`, or `backup`)
- whether you used the Python or Rust implementation
- your OS and tool versions
- a minimal de-identified sample that reproduces the issue, if possible

## FAQ

### Can I share real radiology reports to debug a problem?
No. Please remove patient identifiers and avoid uploading PHI. Synthetic or fully de-identified examples are preferred.

### Which version should I use: Python or Rust?
Use whichever fits your workflow. The repository includes both and `python test_equivalence.py` is intended to check parity between them.

### Why did formatting change after a conversion?
Some conversions intentionally normalize formatting. DOCX and RTF parsing can also depend on style usage in the source file. If output looks wrong, open an issue with a minimal de-identified sample.

### Why did `backup` move files?
`backup` uses `reports_index.json` to decide which files are expected. Regenerate the index before running backup if you recently added templates.

### Where should general questions or setup help go?
Open a GitHub issue for now. If Discussions are enabled later, routine questions can move there.
