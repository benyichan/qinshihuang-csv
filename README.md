<p align="center">
  <img src="assets/opening-zh.png" alt="qinshihuang-csv" width="100%">
</p>

<h1 align="center">秦始皇 CSV 清洗器</h1>

<p align="center">
  <strong>电商数据大一统 — 编码检测 / 脏数据清洗 / 防科学计数法</strong>
</p>

<p align="center">
  <a href="README_EN.md">English</a> ·
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
</p>

---

## 这是什么

电商运营日常收到的 CSV 文件来自不同平台，两个老问题反复折磨人：

1. **编码混乱** — 同一个文件 Excel 打开是乱码，WPS 打开没问题。文件没坏，是解析规则不匹配。
2. **隐藏脏数据** — 金额字段首尾带币种符号、长数字被 Excel 自动转科学计数法（数据永久丢失）、退款列孤零零一个 `-` 其实是零。

本仓库提供一套完整的解决方案，**四种使用方式**适应不同人群。

---

## 项目结构

```
qinshihuang-csv/
├── skills/                                    # 第一层：Hermes Agent Skill
│   └── qinshihuang-csv/                      #   标准 skill 包格式
│       ├── SKILL.md                          #   skill 元数据
│       ├── scripts/
│       │   └── qinshihuang_cleaner.py        #   CLI 清洗脚本
│       └── references/
│           ├── implementation-notes.md        #   实现笔记
│           └── output-format-decision-log.md  #   输出格式决策
├── prompts/
│   └── trae-python-setup-prompt.txt           # 第二层：环境安装引导
├── src/
│   └── 秦始皇CSV清洗器_源码.py                # 第三层：桌面版完整源码
├── assets/
│   ├── opening-zh.png
│   └── opening-en.png
├── .gitignore
├── LICENSE
├── README.md
└── README_EN.md
```

---

## 四种使用方式

### 🎯 不想折腾，直接使用

没编程基础，不想下载任何编程工具，就想把 CSV 文件清洗干净。

去 [Releases](https://github.com/benyichan/qinshihuang-csv/releases) 下载 `qinshihuang-csv-cleaner.exe`，双击打开，选择 CSV 文件，点击开始清洗即可。

### 📖 想学点东西，按指引一步一步来

没编程基础，但有兴趣学。按照 `prompts/trae-python-setup-prompt.txt` 的指引：

1. 下载安装 [Trae 国内版](https://www.trae.com.cn/)（免费）
2. 打开 Trae，选择 SOLO 模式
3. 把 prompt 内容粘贴给 AI
4. AI 会引导你安装 Python、安装依赖库
5. 完成后就可以运行源码了

### 🧑‍💻 有编程基础，想研究或二开

下载 `src/秦始皇CSV清洗器_源码.py`，这是完整的 tkinter GUI 桌面应用，单文件结构清晰。

```bash
pip install pandas charset-normalizer chardet openpyxl
python src/秦始皇CSV清洗器_源码.py
```

想重新打包为 exe：

```bash
pip install pyinstaller
pyinstaller --onefile --windowed "src/秦始皇CSV清洗器_源码.py" --name=qinshihuang-csv-cleaner
```

### 🤖 Hermes Agent 用户

直接安装 skill：

```bash
hermes skills install ./skills/qinshihuang-csv
```

然后加载 skill，告诉 agent：

> 清洗这个 CSV 文件

agent 会引导你完成编码选择、输出格式等操作。

---

## 核心能力

### 四层编码检测

```
BOM头检测 → charset-normalizer → chardet(中文门槛0.2) → 40+编码广度回退
```

### 字段级脏数据清洗

| 特征 | 策略 |
|------|------|
| 首尾币种符号（￥$¥€£...） | 循环 strip 清除 |
| 首尾空格/逗号/分号/制表符 | 循环 strip 清除 |
| 单元格内换行符/回车符 | 直接移除 |
| 金额列纯 `-` / `–` / `—` | 替换为 `"0"` |
| 12+ 位纯数字（订单号等） | 文本格式保护，防科学计数法 |

### 列格式自动推断（互斥优先）

- **文本格式**：12+位纯数字列，CSV 用前导单引号，XLSX 用 `@` 格式
- **日期格式**：匹配 `YYYY-MM-DD HH:MM:SS` 等 5 种模式
- **数值格式**：含小数点的金额列

### 输出

- **CSV**：始终 `utf-8-sig`（带 BOM），否则中文 Windows Excel 按 GBK 解析乱码
- **XLSX**：openpyxl，值类型先转换再设格式

---

## 依赖

| 库 | 用途 |
|----|------|
| pandas | CSV 读取与 DataFrame 操作 |
| charset-normalizer | 首选编码检测 |
| chardet | 兜底编码检测 |
| openpyxl | XLSX 输出 |
| clevercsv（可选） | CSV 方言检测 |

```bash
pip install pandas charset-normalizer chardet openpyxl
```

---

## License

[MIT](LICENSE) © benyichan

---

**⭐ Star 这个仓库，如果它帮你从另一场 CSV 噩梦中解脱出来。**
