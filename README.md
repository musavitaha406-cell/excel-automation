# Text to Excel Converter

A small command-line tool that converts a text file of `key: value` records into a formatted Excel (`.xlsx`) spreadsheet.

Each record is a block of lines, and records are separated by a blank line. Every distinct key becomes a column and every record becomes a row.

## Example

`input.txt`:

```
name: Ali
city: Kashan
major: Translation

name: Sara
city: Tehran
note: met at 10:30
```

```bash
python auto.py input.txt output.xlsx
```

Result: a sheet with the columns `name`, `city`, `major`, `note` and two rows. Fields missing from a record are left empty.

## Features

- Command-line interface with optional arguments (`--sheet`, `--encoding`, `--rtl`)
- Bold header row, frozen header, and auto-fitted column widths
- Right-to-left sheet option for Persian/Arabic data
- Tolerates messy input: malformed lines are skipped with a warning, values may contain colons, and a UTF-8 BOM is handled
- Clear error messages for a missing file, an empty file, a bad encoding, or an output file that is open in Excel
- Unit tests with the standard library `unittest`

## Installation

```bash
pip install -r requirements.txt
```

Python 3.9+ is required.

## Usage

```bash
python auto.py [INPUT] [OUTPUT] [--sheet NAME] [--encoding ENC] [--rtl]
```

| Argument | Default | Description |
| --- | --- | --- |
| `INPUT` | `input.txt` | Text file to read |
| `OUTPUT` | `output.xlsx` | Excel file to create |
| `--sheet` | `Data` | Sheet name (max 31 characters) |
| `--encoding` | `utf-8` | Input file encoding |
| `--rtl` | off | Make the sheet right-to-left |

## Tests

```bash
python -m unittest -v
```

## Tech Stack

Python, pandas, openpyxl
