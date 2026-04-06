# Repository Guidelines

## Project Structure & Module Organization
Use the `src/` layout. Package code lives in `src/excel_fillna/`: keep workbook logic in `core.py`, CLI argument handling in `cli.py`, and public exports in `__init__.py`. Tests live in `tests/`, with binary workbook fixtures under `tests/fixtures/`. Specs and behavior notes belong in `docs/` (`FUNCTIONAL_SPEC.md`, `CLI_SPEC.md`). Treat `build/` as generated output, not hand-edited source.

## Build, Test, and Development Commands
Create or refresh a local environment with `python3 -m venv .venv` and install dependencies with `python3 -m pip install -e ".[dev]"`. Install the package normally with `python3 -m pip install .` when you want the `fna` console script available outside editable mode. Run the test suite with `python3 -m pytest`; this is the verified command for the current workspace. For quick manual checks, use the CLI directly, for example `fna workbook.xlsx --range A1:C20 --fill-text MISSING`.

## Coding Style & Naming Conventions
Follow existing Python style: 4-space indentation, type hints on public functions, dataclasses for structured results, and `snake_case` for functions, variables, and test names. Keep module-level constants in `UPPER_SNAKE_CASE`. Favor small helpers over deeply nested logic, especially in worksheet and XML patching paths. No formatter or linter is configured in `pyproject.toml`, so match the surrounding code closely and keep imports, docstrings, and argument wrapping consistent with existing files.

## Testing Guidelines
Write tests with `pytest` in files named `test_*.py` and functions named `test_*`. Cover both library behavior and CLI behavior; current examples are in `tests/test_core.py` and `tests/test_cli.py`. Use `tmp_path` for output workbooks and preserve fixture-based checks for images, charts, and related archive entries when changing save logic.

## Commit & Pull Request Guidelines
Current history uses Conventional Commit style (`feat: create excel fillna library`), so prefer prefixes like `feat:`, `fix:`, and `test:`. Keep commit subjects imperative and concise. PRs should summarize the workbook scenario changed, list the commands run (`python3 -m pytest`), and note any fixture updates or behavioral changes to merge handling, exclusions, or output naming.

## Fixture & Workbook Integrity
This project intentionally patches worksheet XML to avoid losing workbook artifacts. Do not re-save fixture files casually through Excel or `openpyxl`; update them only when the artifact set is intentionally changing, and mention that explicitly in the PR.
