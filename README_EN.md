<p align="center">
  <img src="assets/opening-en.png" alt="qinshihuang-csv" width="100%">
</p>

<h1 align="center">qinshihuang-csv</h1>

<p align="center">
  <strong>E-commerce CSV encoding detection & data cleansing toolkit</strong>
</p>

<p align="center">
  <a href="README.md">中文</a> · <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
</p>

---

## What is this

CSV files exported from Chinese e-commerce platforms have two common problems:

1. **Encoding inconsistency**. The same file opens as gibberish in Excel but fine in WPS. The file isn't broken — the parsing rules just don't match.
2. **Hidden dirty data**. Leading/trailing spaces, currency symbols embedded in price fields, long numbers silently converted to scientific notation (data loss, permanent), and lone `-` characters in refund columns that actually mean zero.

This skill detects the encoding of any CSV file, cleans up dirty fields, and outputs standardized CSV/XLSX files.

## Stack

| Dependency | Purpose |
|------------|---------|
| Python 3.13+ | Runtime |
| pandas | CSV I/O, DataFrame handling |
| charset-normalizer | Primary encoding detection |
| chardet | Fallback encoding detection |
| openpyxl | XLSX output (Excel 2016+) |
| clevercsv (optional) | CSV dialect detection (delimiter/quoting) |

## Quick Start

```bash
pip install pandas charset-normalizer chardet openpyxl

# Single file
python scripts/qinshihuang_cleaner.py your_file.csv --to-xlsx

# Batch directory
python scripts/qinshihuang_cleaner.py your_dir/ --batch --to-xlsx
```

Hermes Agent users: load the skill and tell the agent to clean a CSV.

## Core Capabilities

### 4-Layer Encoding Detection

```
BOM Detection → charset-normalizer → chardet (0.2 threshold for Chinese) → 40+ Encoding Fallback
```

### 3-Step Cleaning

```
① Auto-detect header → ② Per-column dirty pattern analysis → ③ Full clean + output
```

### Column Format Auto-Detection (mutually exclusive)

```
Text (@, ≥12 digits) > Date (YYYY-MM-DD HH:MM:SS) > Numeric (#,##0.00)
```

## Pitfalls

### Encoding Detection

chardet alone isn't enough. Chinese-encoded files can score as low as 20.7% confidence, and a single corrupted byte can break all GBK/GB18030 decoding, making the fallback land on cp1252 with garbled output. The fix: a 4-layer pipeline plus CJK-tolerant fallback that decodes with `errors='replace'` and validates by Chinese character count.

### XLSX Format ≠ Type Conversion

openpyxl's `number_format` only changes display, not the underlying data type. Writing a string `"2025-10-19"` and setting a date format still leaves it as text in Excel. Values must be converted first (`datetime.strptime` → datetime, `float()` → numeric), then formatted.

### Long Number Precision Loss

Numbers over 12 digits get converted to scientific notation by Excel — data loss is permanent. The fix: a leading single quote `'123456789012345` in CSV output. Excel hides the quote and displays the pure number. Threshold is 12 digits (not 15) because product IDs at 12 digits can also trigger scientific notation.

### UTF-8 Without BOM

Chinese Windows Excel interprets BOM-less UTF-8 as the local encoding (GBK), producing garbage. CSV output always uses `utf-8-sig` regardless of user's encoding choice.

### Dash Placeholder

A lone `-` in amount columns means zero, not missing data. Automatically detected and replaced with `"0"`.

## Known Issues

- **Large file performance**: Cell-by-cell iteration slows above 100K rows. Vectorized approach planned.
- **chardet CJK confidence**: Chinese-encoded files may score as low as 0.2 — handled with a lowered threshold.
- **Windows installation**: Some dependencies (clevercsv) may fail due to SSL certificate issues — core functionality unaffected.

---

**⭐ Star this repo if it saved you from another CSV headache.**

[MIT](LICENSE) © benyichan
