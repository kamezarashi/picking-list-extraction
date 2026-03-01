# AGENTS Guide for `picking-list-extraction`

This file defines conventions for coding agents working in this repository.

## 1) Project Snapshot

- Language: Python
- Entrypoint: `generate_picking_list.py`
- Goal: read order CSVs and output formatted Excel picking lists
- Input contract: `input/*.csv`
- Output contract: `output/<csv_name>/1_商品別.xlsx`, `2_センター別.xlsx`, `3_店舗別.xlsx`
- Current layout: single-script project (no package split yet)

## 2) Environment Setup

Use a local virtual environment before running commands.

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
pip install pandas openpyxl
```

If dependency files are added, prefer them:

```bash
pip install -r requirements.txt
```

## 3) Build / Run Commands

There is no compile step. Treat "build" as "script runs end-to-end".

```bash
python generate_picking_list.py
```

Expected behavior:
- scans `input/*.csv`
- creates subfolders in `output/`
- writes 3 Excel files per source CSV

## 4) Lint / Format Commands

No lint/format config is committed yet. Use these defaults:

```bash
ruff check .
ruff check . --fix
black .
```

If Ruff/Black are unavailable, report that clearly.

## 5) Test Commands (Including Single Test)

No tests are currently committed, but new tests should use `pytest`.

Run all tests:

```bash
pytest
```

Run one file:

```bash
pytest tests/test_generate_picking_list.py
```

Run one test function (node id):

```bash
pytest tests/test_generate_picking_list.py::test_normalize_size_text
```

Run by keyword:

```bash
pytest -k "normalize_size"
```

If `unittest` is used instead, run one test with:

```bash
python -m unittest tests.test_generate_picking_list.TestName.test_method
```

## 6) Verification Flow for Code Changes

Use this order unless task constraints require otherwise:

1. Run targeted tests for changed logic.
2. Run full test suite (if present).
3. Run lint/format checks.
4. Run `python generate_picking_list.py` with representative input if processing logic changed.

Always report what you ran and what could not be run.

## 7) Code Style Guidelines

### Imports

- Order imports: standard library, third-party, local modules.
- Prefer one import per line where practical.
- Remove unused imports.
- Avoid function-local imports except for optional/error-only paths.

### Formatting

- Follow PEP 8.
- Use 4 spaces (no tabs).
- Keep lines readable (target <= 88 chars).
- Prefer helper functions over deeply nested blocks.
- Add comments only where intent is not obvious.

### Types and Data Contracts

- Add type hints for new/modified functions.
- Use explicit return types (`-> None`, `-> pd.DataFrame`, etc.).
- Document required DataFrame columns in docstrings for transformation helpers.
- Validate assumptions early (columns, row indexes, encoding).

### Naming

- `snake_case` for variables/functions.
- `UPPER_SNAKE_CASE` for constants.
- Prefer descriptive names over abbreviations (except domain terms like `JAN`, `MK`).
- Preserve Japanese column labels that are part of I/O contracts.

### Error Handling and Logging

- Fail fast for invalid file/schema structure.
- Skip problematic input files gracefully when possible.
- Prefer narrow exceptions over broad `except Exception`.
- Include actionable context in errors (file, phase, column/index).
- Do not silently swallow errors.

### I/O and Side Effects

- Keep transformation logic separate from file read/write when practical.
- Avoid machine-specific absolute paths.
- Preserve `input/` and `output/` contracts unless migration is explicitly requested.

### Pandas / Excel Practices

- Prefer vectorized operations over row loops where feasible.
- Use `.copy()` intentionally to avoid view/copy ambiguity.
- Normalize string fields before grouping and sorting.
- Keep sheet-name sanitization and 31-char truncation centralized.
- Keep `openpyxl` formatting concerns separated from aggregation logic.

## 8) Repo-Specific Notes

- `generate_picking_list.py` is the primary implementation; keep changes incremental.
- Maintain compatibility with the current column-index-based CSV extraction.
- If you add tool config (`pyproject.toml`, `pytest.ini`, etc.), keep it minimal and update this file.

## 9) Cursor / Copilot Rules

Checked locations:
- `.cursor/rules/`
- `.cursorrules`
- `.github/copilot-instructions.md`

Current status: no Cursor or Copilot rule files exist in this repository.
If these are added later, update this section and follow them.
