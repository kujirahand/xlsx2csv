# xlsx2csv

A CLI tool to convert Excel files into CSV files, one per sheet.  
Supports xlsx, xls, ods, and more.

> 日本語の説明: [README-ja.md](README-ja.md)

## Features

- Simple command-line interface for quick conversion of Excel files to CSV
- Rust implementation for fast performance and low memory usage
- Converts each sheet in an Excel file to a separate CSV file
- Supports multiple output encodings (Shift-JIS by default, UTF-8, EUC-JP, etc.)
- Date cells are output in `YYYY-MM-DD` format by default
- Fields containing commas, double quotes, or newlines are escaped per RFC 4180
- Output files are placed in the same directory as the input file
- On Windows, you can drag and drop an Excel file onto the executable to convert it

## Download

Download the latest binary from [Releases](https://github.com/kujirahand/xlsx2csv/releases/).

## Configuration

Place an `xlsx2csv.toml` file in the same directory as the executable or in the current working directory to change default behavior.

```toml
[default]
encoding = "utf-8"       # output encoding (default: shift_jis)
date_format = "%Y/%m/%d" # date format (default: %Y-%m-%d)
output_format = "tsv"    # output format: csv or tsv (default: csv)
```

### date_format examples

| Value | Output |
|---|---|
| `%Y-%m-%d` | 2025-01-31 |
| `%Y/%m/%d` | 2025/01/31 |
| `%d/%m/%Y` | 31/01/2025 |
| `%Y年%m月%d日` | 2025年1月31日 |

Format specifiers follow [chrono's strftime syntax](https://docs.rs/chrono/latest/chrono/format/strftime/index.html).

## Command-line Options

Command-line options override values from the configuration file.

| Option | Config key | Default | Description |
|---|---|---|---|
| `-e`, `--encoding` | `encoding` | `shift_jis` | Output character encoding |
| `-d`, `--date_format` | `date_format` | `%Y-%m-%d` | Date format |
| `--tsv` | `output_format` | `csv` | Output as tab-separated (TSV) |

## Building (for developers)

Requires the Rust toolchain.

```bash
cargo build --release
```

The binary is generated at `target/release/xlsx2csv`.

## Usage

```bash
xlsx2csv [-e <encoding>] [-d <date_format>] [--tsv] [--utf8] <excel_file>
```

### Options

| Option | Description | Default |
|---|---|---|
| `-e`, `--encoding` | Output encoding | `shift_jis` |
| `-d`, `--date_format` | Date format | `%Y-%m-%d` |
| `--tsv` | Output as TSV (tab-separated) | (off) |
| `-u8`, `--utf8` | Output as UTF-8 (shortcut for `--encoding utf-8`) | (off) |
| `-h`, `--help` | Show help | - |

### Examples

```bash
# Output as Shift-JIS (default)
xlsx2csv data.xlsx

# Output as UTF-8
xlsx2csv -e utf-8 data.xlsx

# Output as EUC-JP
xlsx2csv --encoding euc-jp data.xlsx

# Output dates in YYYY/MM/DD format
xlsx2csv -d "%Y/%m/%d" data.xlsx

# Specify encoding and date format together
xlsx2csv -e utf-8 -d "%Y/%m/%d" data.xlsx

# Output as TSV
xlsx2csv --tsv data.xlsx

# Output as TSV with UTF-8 encoding
xlsx2csv --tsv --utf8 data.xlsx
```

Output files are named `<filename>_<sheetname>.csv` (or `.tsv`) per sheet.

```
data_Sheet1.csv
data_Sheet2.csv
```

## Specification

- Supported formats: xlsx, xls, ods, etc. (auto-detected from extension)
- Output directory: same as the input file
- File naming: `<input_filename>_<sheet_name>.csv`
- Output encoding: Shift-JIS by default (changeable with `-e utf-8`, etc.)
- Date format: `%Y-%m-%d` by default (changeable with `-d "%Y/%m/%d"`, etc.)
- Fields with commas, double quotes, or newlines are escaped per RFC 4180

## License

MIT License
