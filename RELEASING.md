# Releasing

This repository is not ready for a stable `v1.0.0` yet.

Until the criteria below are met, mark GitHub releases as **pre-release** and use `v0.y.z` tags.

## Versioning policy

- Use semantic versioning tags prefixed with `v`, for example `v0.1.0`.
- While the project is pre-1.0, breaking changes, template reshuffles, and converter behavior changes may still happen between minor versions.
- Mark GitHub releases as **pre-release** until the first stable release gate is satisfied.

## Release checklist

1. Update `CHANGELOG.md` from `Unreleased` with the user-visible changes.
2. Set up Python dependencies before verification. Create and activate a virtual environment, or run an explicit dependency install in an environment that allows package installs; use the [QUICKSTART.md installation steps](QUICKSTART.md#install-python-deps) for the exact commands.
3. From that clean checkout with the Python environment active, confirm the commands below still work:
   - `python run.py --help`
   - `python test_equivalence.py`
   - `cd rust_converters && cargo build`
4. Make sure installation and usage notes in `README.md` and `QUICKSTART.md` still match reality, especially the dependency setup path.
5. Verify no sample, fixture, screenshot, or generated artifact contains PHI or private medical data.
6. Draft or review the GitHub release notes before publishing.

## Gate for the first stable release

Do not publish `v1.0.0` until all of the following are true:

- The Python and Rust workflows have repeatable validation on the canonical template corpus.
- Installation and build steps have been tested from scratch and documented without guesswork.
- The supported conversion workflows and known limitations are documented clearly enough for external users.
- At least one pre-release has been exercised to shake out release-note quality and upgrade guidance.

If those conditions are not met, publish the next change as a pre-release instead of a stable release.
