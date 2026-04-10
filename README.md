# Data Curation Workflows
authors: Ellery Galvin\
date: 2026-03-24\
contact: ellery.galvin@colorado.edu, crdds@colorado.edu

## Description
This project checks tabular datasets (`.csv` and `.xlsx`) for common curation issues.  
Each test reports:
- what failed,
- where it failed, and
- an example offending value.

## Installation

```bash
pip install FAIRspreadsheets
```

## Usage
```python
import fair
suite = fair.Test_Suite("your_data.csv")
suite.run()
suite.report()
suite.save(format = "csv")
```

## Call signatures
Test_Suite(wb_path, to_run=None, to_skip=None, config_path=None)
- wb_path: path to the workbook

Which tests to run?  Defaults to all.  Else, specify one of
- to_run: list of test names to run, lower case strings as below
- to_skip: list of test names to skip, lower case strings as below

Extra arguments to replace defaults for tests (see below):
- config_path: path to the config file, replaces defaults where specified

run(): runs the tests
report(): prints the results to the console
save(format = "json"): saves the results to a JSON file
save(format = "csv"): saves the results to a CSV file


## Tests Implemented

### File-level tests
- `file_name_length`: filename length is within allowed bounds
- `file_name_whitespace`: filename contains no whitespace
- `file_name_final`: filename does not include the word `final`
- `file_name_word_separation`: filename does not mix camelCase, underscores, and dashes
- `file_name_special_characters`: filename has no disallowed special characters
- `file_encoding`: text files decode cleanly in the configured encoding

### Sheet-level tests
- `sheet_empty`: worksheet is not empty
- `sheet_name`: sheet name contains no whitespace or disallowed special characters
- `sheet_upper_left_corner`: table begins at upper-left (no leading blank rows/columns)
- `sheet_multi_table`: sheet does not appear to contain multiple separate tables

### Header-level tests
- `header_duplicates`: duplicate header names are flagged
- `header_id`: first header should not be exactly `ID`
- `header_length`: header length is within allowed bounds
- `header_first_char`: headers should not start with digits
- `header_space`: headers should not contain spaces
- `header_word_separation`: headers should not mix camelCase, underscores, and dashes
- `header_special_characters`: headers should not contain disallowed special characters
- `header_date`: date-like header names (for combined date fields) are flagged
- `header_mixed_datatypes`: columns with mixed numeric/text values are flagged

### Cell-level tests
- `cell_aggregate_row`: aggregate-like summary row at bottom (e.g., total/sum) is flagged
- `cell_special_characters`: disallowed special characters in cells are flagged
- `cell_untrimmed_white_space`: leading/trailing whitespace is flagged
- `cell_newlines_tabs`: tabs/newlines/vertical tabs in cells are flagged
- `cell_missing_value_text`: text placeholders for missing values (`NA`, `null`, `-`, etc.) are flagged
- `cell_question_mark_only`: cells equal to `?` are flagged
- `cell_white_space_only`: whitespace-only cells are flagged
- `cell_number_space`: cells containing only digits+spaces are flagged
- `cell_dates`: dates not matching required format are flagged
- `cell_scientific_notation`: scientific notation in text cells is flagged
- `cell_units`: values that combine number+unit in one cell (e.g., `10kg`) are flagged


## Options Configuration
Options are set in a JSON config file (see `config.example.json`).

Important conventions:
- Top-level key = lower-case class name (for example `cell_dates`)
- Option names = keyword-only `__init__` args for that test
- Only the tests and options below are accepted

### File-level options
- `file_name_length`
  - `min_length` (int, default `5`): minimum allowed filename length
  - `max_length` (int, default `32`): maximum allowed filename length
- `file_name_special_characters`
  - `special_char_pattern` (regex string, default `[!@#$%^&*()+=\[\]{};:'"|\\,<>\?/]`): disallowed characters
- `file_encoding`
  - `valid_encoding` (string, default `utf-8`): expected text encoding for CSV-like files

### Sheet-level options
- `sheet_name`
  - `special_char_pattern` (regex string, default `[!@#$%^&*()+=\[\]{};:'"|\\,<>\?/]`): disallowed characters

### Header-level options
- `header_length`
  - `min_length` (int, default `1`): minimum header length
  - `max_length` (int, default `24`): maximum header length
- `header_special_characters`
  - `special_char_pattern` (regex string, default `[!@#$%^&*()+=\[\]{};:'"|\\,<>\?/]`): disallowed characters
- `header_date`
  - `date_keywords` (list of strings, default `["date", "datetime", "timestamp"]`): keywords used to detect date-like headers

### Cell-level options
- `cell_aggregate_row`
  - `aggregate_words` (list of strings, default `["total","sum","average","count","min","max"]`): terms used to detect summary rows
- `cell_special_characters`
  - `default_pattern` (regex string): disallowed chars for normal columns
  - `url_pattern` (regex string, default `^https?://`): validation pattern for URL columns
  - `free_text_pattern` (regex string): disallowed chars for free-text columns
  - `url_columns` (list of strings, default `[]`): columns validated with `url_pattern`
  - `free_text_columns` (list of strings, default `[]`): columns validated with `free_text_pattern`
  - `skip_columns` (list of strings, default `[]`): columns skipped by this test
- `cell_dates`
  - `date_columns` (list of strings, default `[]`): explicit date columns
  - `auto_detect_columns` (bool, default `false`): infer likely date columns automatically
  - `format_code` (string, default `%Y/%m/%d`): required `datetime.strptime` date format
  - `date_column_threshold` (float `0..1`, default `0.8`): fraction of parseable values needed to auto-mark a column as date-like
- `cell_units`
  - `unit_abbreviations` (list of strings): unit suffixes used to detect values like `12kg`

Tests without options should be configured with an empty object or omitted.

## Output Specification
Only failed (or not-completed) tests appear in saved output. Passing tests are omitted.

### JSON output (`results/<input_base>_<YYYY-MM-DD_HH-MM-SS>.json`)
- Root object: keys are sheet names, plus `file` for file-level failures
- Value per sheet: object keyed by test name
- Value per test:
  - `message`: failure/not-completed message
  - `issues`: object of `location -> example`
- For JSON compatibility, non-string issue keys (such as tuple coordinates) are stringified.

### CSV output (`results/<input_base>_<YYYY-MM-DD_HH-MM-SS>.csv`)
Columns are exactly:
- `path`: input workbook path
- `sheet`: sheet name (`file` for file-level tests)
- `test_name`: lower-case test class name
- `message`: failure/not-completed message
- `location`: issue key (string form)
- `example`: issue value

Each row corresponds to one issue entry from one failed test.

## For Developers


## Directory contents
- .gitignore excludes the data directory from version control, and other files
- config.example.json is an example configuration for test options
- demo.xlsx is a file on which to test detectors.py
- detectors.py is the main module that implements a series of automated tests for data curation workflows
- pyproject.toml is the specification for the PyPI package
- LICENSE contains the licensing information for this software


### Architecture overview
- `Test_Suite` discovers concrete tests automatically from subclasses of `Test`
- Tests are grouped by level using class name prefixes: `file_`, `sheet_`, `header_`, `cell_`
- Dependencies are declared via `Has_Dependency`; the suite executes tests in dependency-safe order
- Each test writes failures into `self.issues`, then calls `Test.validate(...)` to set status/message
- `trimmed_results()` filters outputs to failed/not-completed tests only

### How to add a new test safely
1. Choose the right base class:
   - file checks: inherit `File`
   - sheet checks: inherit `Sheet`
   - header checks: inherit `Header`
   - cell checks: inherit `Cell`
2. Name the class with underscores (example `Cell_My_New_Check`).  
   The runtime test key becomes lower-case class name (`cell_my_new_check`).
3. Put user-configurable settings in keyword-only args:
   - use `def __init__(self, *, ...)`
   - validate types/ranges in `__init__`
4. Implement `validate(...)` and populate `self.issues` as:
   - key: location (header name, tuple coordinate, `filename`, etc.)
   - value: offending example/value
5. End validation with `Test.validate(self, fail_message, pass_message)` so status/message are consistent.
6. If your test depends on another test, add `Has_Dependency.__init__(self, DependencyClass)` and use dependency outputs in `validate`.
7. Add documentation:
   - include the new test in the test list above
   - add config options under the correct section (if any)
   - update `config.example.json` with sensible defaults
