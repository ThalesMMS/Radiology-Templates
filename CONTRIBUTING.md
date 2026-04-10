# Contributing

Thanks for improving Radiology-Templates.

This repository is a practical collection of radiology report templates plus Python/Rust conversion tools. Keep contributions small, reviewable, and easy to validate.

## Good first contributions

- Fix template formatting or wording issues
- Add missing README or quickstart clarifications
- Improve conversion reliability for real-world template files
- Add tests for edge cases in DOCX/Markdown/TXT round-trips
- Report and document compatibility differences between the Python and Rust implementations

## Before you open a pull request

1. Open an issue first for large changes or new workflows.
2. Keep medical content de-identified. Do not include PHI, screenshots with patient data, or real reports.
3. Prefer focused PRs over broad cleanup.
4. Update docs when behavior or file layout changes.

## Local checks

From the repository root:

```bash
python test_equivalence.py
```

If you are working on the Rust side, also run:

```bash
cd rust_converters
cargo build
```

If you are working on the Python side, make sure the entrypoint still works:

```bash
python run.py --help
```

## Style notes

- Preserve the existing folder structure unless there is a strong reason to change it.
- Keep template conversions deterministic when possible.
- Document any formatting assumptions that affect output.
- Avoid introducing dependencies unless they clearly reduce maintenance cost.

## Pull request checklist

- [ ] Change is scoped and explained clearly
- [ ] No PHI or private medical data was added
- [ ] Docs were updated if needed
- [ ] Validation steps and results are included in the PR description
- [ ] Related issue is linked when applicable

## Code of conduct

By participating, you agree to follow the [Code of Conduct](CODE_OF_CONDUCT.md).
