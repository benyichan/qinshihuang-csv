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

CSV files from Chinese e-commerce platforms have two common problems:

1. **Encoding inconsistency** — Gibberish in Excel, fine in WPS. The file isn't broken, the parsing rules just don't match.
2. **Hidden dirty data** — Currency symbols in price fields, long numbers silently converted to scientific notation (permanent data loss), lone `-` in refund columns that actually means zero.

**Four ways to use** for different levels of users.

---

## Project Structure

```
qinshihuang-csv/
├── skills/
│   └── qinshihuang-csv/               # Layer 1: Standard Hermes Skill
│       ├── SKILL.md
│       ├── scripts/
│       │   └── qinshihuang_cleaner.py
│       └── references/
├── prompts/
│   └── trae-python-setup-prompt.txt    # Layer 2: Trae setup guide
├── src/
│   └── 秦始皇CSV清洗器_源码.py         # Layer 3: Desktop app source
├── assets/
├── LICENSE
├── README.md
└── README_EN.md
```

---

## Four Ways to Use

### 🎯 Just want to clean CSV files

No programming background, no time to learn tools.  
Download `qinshihuang-csv-cleaner.exe` from [Releases](https://github.com/benyichan/qinshihuang-csv/releases). Double-click to run, select your CSV files, click start.

### 📖 Want to learn along the way

No programming background but interested. Follow `prompts/trae-python-setup-prompt.txt`:

1. Install [Trae CN](https://www.trae.com.cn/) (free)
2. Open Trae, select SOLO mode
3. Paste the prompt to AI
4. AI guides you through Python and dependency setup
5. You can then run the source code

### 🧑‍💻 Developer — study or modify

Download `src/秦始皇CSV清洗器_源码.py` — a clean single-file tkinter desktop app.

```bash
pip install pandas charset-normalizer chardet openpyxl
python src/秦始皇CSV清洗器_源码.py
```

Repackage to exe:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed "src/秦始皇CSV清洗器_源码.py" --name=qinshihuang-csv-cleaner
```

### 🤖 Hermes Agent user

```bash
hermes skills install ./skills/qinshihuang-csv
```

Load the skill and tell the agent: clean this CSV file.

---

## Core Capabilities

### 4-Layer Encoding Detection

```
BOM Detection → charset-normalizer → chardet (0.2 threshold for Chinese) → 40+ Encoding Fallback
```

### Data Cleaning

- Strip leading/trailing currency symbols, spaces, commas, tabs
- Remove internal newlines and carriage returns
- Replace lone `-`/`–`/`—` in amount columns with `"0"`
- Protect 12+ digit numbers from Excel scientific notation

### Column Format Auto-Detection

Text format (12+ digits) > Date format (YYYY-MM-DD HH:MM:SS) > Numeric format (#,##0.00)

### Output

- **CSV**: Always `utf-8-sig` (BOM required for Chinese Excel compatibility)
- **XLSX**: openpyxl, type conversion before formatting

## Dependencies

```bash
pip install pandas charset-normalizer chardet openpyxl
```

## License

[MIT](LICENSE) © benyichan

---

**⭐ Star this repo if it saved you from another CSV headache.**
