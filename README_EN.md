<p align="center">
  <img src="assets/opening-en.png" alt="qinshihuang-csv" width="100%">
</p>

<h1 align="center">qinshihuang-csv</h1>

<p align="center">
  <strong>E-commerce CSV encoding detection & data cleansing toolkit</strong>
</p>

<p align="center">
  <a href="README.md">中文</a> ·
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
</p>

---

## What is this

CSV files exported from Chinese e-commerce platforms have two common problems:

1. **Encoding inconsistency**. The same file opens as gibberish in Excel but fine in WPS. The file isn't broken — the parsing rules just don't match.
2. **Hidden dirty data**. Leading/trailing spaces, currency symbols embedded in price fields, long numbers silently converted to scientific notation (data loss, permanent), and lone `-` characters in refund columns that actually mean zero.

This repository packages a **three-layer solution** covering everyone from absolute beginners to developers.

## Three Layers

```
Layer 1: Skill — Hermes Agent skill for CSV cleaning
         scripts/qinshihuang_cleaner.py
         For: Hermes Agent users

Layer 2: Prompt — Trae environment setup guide
         prompts/trae-python-setup-prompt.txt
         For: Zero-experience users (install Trae → paste prompt → AI guides)

Layer 3: Source + exe — Standalone desktop GUI app
         src/ + dist/
         For: Developers who want to study or modify the code
```

## Quick Start

### 🎯 Zero-experience users

1. Download [Trae CN](https://www.trae.com.cn/) (free, no payment required)
2. Open Trae, select **SOLO mode**
3. Copy the content of `prompts/trae-python-setup-prompt.txt` and paste to AI
4. AI guides you through installing Python and dependencies
5. Double-click `dist/秦始皇CSV清洗器.exe` to use

### 🛠 Python users

```bash
pip install pandas charset-normalizer chardet openpyxl
python src/秦始皇CSV清洗器_源码.py
```

### 🧑‍💻 Hermes Agent users

Load the `qinshihuang-csv` skill and tell the agent to clean a CSV file.

### 📦 Repackaging

```bash
pip install pyinstaller
pyinstaller --onefile --windowed "src/秦始皇CSV清洗器_源码.py" --name=秦始皇CSV清洗器
```

The new `.exe` will be at `dist/`.

## Project Structure

```
qinshihuang-csv/
├── prompts/
│   └── trae-python-setup-prompt.txt   # Layer 2: Trae setup guide
├── src/
│   └── 秦始皇CSV清洗器_源码.py         # Layer 3: Full GUI desktop app (tkinter)
├── dist/
│   └── 秦始皇CSV清洗器.exe             # Layer 3: Pre-built single-file exe
├── scripts/
│   └── qinshihuang_cleaner.py         # Layer 1: Hermes Agent CLI version
├── references/
│   ├── implementation-notes.md
│   └── output-format-decision-log.md
├── assets/
│   ├── opening-zh.png
│   └── opening-en.png
├── README.md
├── README_EN.md
└── LICENSE
```

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

- **chardet CJK confidence**: Chinese-encoded files may score as low as 0.2 — handled with a lowered threshold.
- **XLSX format ≠ type conversion**: openpyxl's `number_format` only changes display. Values must be converted first (`strptime` → datetime, `float()` → numeric).
- **Long number precision loss**: Numbers ≥12 digits trigger scientific notation in Excel. Fixed with leading single-quote text format.
- **UTF-8 without BOM**: Chinese Windows Excel interprets BOM-less UTF-8 as GBK. CSV output always uses `utf-8-sig`.
- **Dash placeholder**: `-` in amount columns means zero, not missing data. Automatically replaced with `"0"`.
- **Large file performance**: Cell-by-cell iteration slows above 100K rows. Vectorized optimization planned.

## Stack

| Dependency | Purpose |
|------------|---------|
| Python 3.10+ | Runtime |
| pandas | CSV I/O, DataFrame handling |
| charset-normalizer | Primary encoding detection |
| chardet | Fallback encoding detection |
| openpyxl | XLSX output (Excel 2016+) |
| clevercsv (optional) | CSV dialect detection (delimiter/quoting) |

## License

[MIT](LICENSE) © benyichan

---

**⭐ Star this repo if it saved you from another CSV headache.**
