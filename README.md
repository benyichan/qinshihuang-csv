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

本仓库一次打包了**三层解决方案**，覆盖从纯小白到开发者的完整链路。

## 三层架构

```
┌─ 第一层: Skill ─────────────────────────────────────────┐
│  Hermes Agent 可加载的清洗 Skill                         │
│  面向: 已有 Hermes Agent 的用户                          │
│  位置: scripts/qinshihuang_cleaner.py                    │
│  用法: 加载 skill 后告诉 agent "清洗这个 CSV"             │
├─ 第二层: Prompt ─────────────────────────────────────────┤
│  Trae 环境安装引导 Prompt                                │
│  面向: 零基础普通人                                      │
│  位置: prompts/trae-python-setup-prompt.txt              │
│  用法: 安装 Trae → 粘贴 prompt → AI 自动引导             │
├─ 第三层: 源码 + exe ─────────────────────────────────────┤
│  可直接运行的桌面应用                                    │
│  面向: 想研究/修改代码的开发者                           │
│  位置: src/ + dist/                                      │
│  用法: 双击 exe 或 python 运行源码                       │
└──────────────────────────────────────────────────────────┘
```

---

## 快速上手

### 🎯 零基础用户（推荐）

1. 下载安装 [Trae 国内版](https://www.trae.com.cn/)（免费，无需付费）
2. 打开 Trae，选择 **SOLO 模式**
3. 复制 `prompts/trae-python-setup-prompt.txt` 的全部内容，粘贴给 AI
4. AI 会一步步引导你安装 Python、安装依赖库
5. 完成后双击 `dist/秦始皇CSV清洗器.exe` 即可使用

### 🛠 有 Python 基础的用户

```bash
pip install pandas charset-normalizer chardet openpyxl
python src/秦始皇CSV清洗器_源码.py
```

### 🧑‍💻 Hermes Agent 用户

加载 `qinshihuang-csv` skill 后，告诉 agent：

> 清洗这个 CSV 文件

### 📦 想研究源码二次打包

```bash
pip install pyinstaller
pyinstaller --onefile --windowed "src/秦始皇CSV清洗器_源码.py" --name=秦始皇CSV清洗器
```

打包后的 `.exe` 在 `dist/` 目录下。

---

## 项目结构

```
qinshihuang-csv/
├── prompts/
│   └── trae-python-setup-prompt.txt   # 第二层：Trae 环境安装引导
├── src/
│   ├── qinshihuang/
│   │   └── 秦始皇CSV清洗器_源码.py     # 第三层：完整 GUI 桌面应用（tkinter）
├── dist/
│   └── 秦始皇CSV清洗器.exe             # 第三层：预编译单文件 exe
├── scripts/
│   └── qinshihuang_cleaner.py         # 第一层：Hermes Agent CLI 版
├── references/
│   ├── implementation-notes.md        # 清洗实现笔记
│   └── output-format-decision-log.md  # 输出格式决策记录
├── assets/
│   ├── opening-zh.png                 # 中文版横幅
│   └── opening-en.png                 # 英文版横幅
├── README.md                          # 中文说明
├── README_EN.md                       # 英文说明
└── LICENSE                            # MIT 协议
```

---

## 核心能力

### 四层编码检测

```
BOM头检测 → charset-normalizer → chardet(中文门槛0.2) → 40+编码广度回退
```

中文 CSV 的编码检测不能只靠 chardet。单个坏字节就能让 GBK 解码失败，回退落到 cp1252 输出中文乱码。四层流水线 + CJK 容错回退（`errors='replace'` + 汉字数验证）解决了这个问题。

### 字段级脏数据清洗

| 特征 | 策略 |
|------|------|
| 首尾币种符号（￥$¥€£...） | 循环 strip 清除 |
| 首尾空格/逗号/分号/制表符 | 循环 strip 清除 |
| 单元格内换行符/回车符 | 直接移除 |
| 金额列纯 `-` / `–` / `—` | 替换为 `"0"` |
| 12+ 位纯数字（订单号等） | 文本格式保护，防科学计数法 |

### 列格式自动推断（互斥优先）

```
文本格式 (@) > 日期格式 (YYYY-MM-DD HH:MM:SS) > 数值格式 (#,##0.00)
```

- **文本格式**：12+位纯数字列，CSV 用前导单引号 `'123...`，XLSX 用 `@` 格式
- **日期格式**：匹配 `YYYY-MM-DD HH:MM:SS` 等 5 种模式，值转 `datetime` 对象再写入
- **数值格式**：含小数点的金额列，值转 `float` 再写入

### 输出

- **CSV**：始终使用 `utf-8-sig`（带 BOM），否则中文 Windows Excel 按 GBK 解析 -> 乱码
- **XLSX**：openpyxl，值类型必须先转换（datetime/float），再设格式，否则 Excel 仍按文本处理

---

## 踩坑记录

### 编码检测

chardet 对中文文件的置信度可能低至 0.2，但 `language='zh'` 字段依然正确。检测时对中文降门槛到 0.2。

### XLSX 格式 ≠ 类型转换

openpyxl 的 `number_format` 只改变显示方式。字符串 `"2025-10-19"` 设日期格式后仍是文本。必须先用 `strptime` 转 `datetime` 对象，再设格式。

### 长数字精度丢失

Excel 15 位有效数字限制，12 位以上的订单号已可能触发科学计数法。CSV 用前导单引号防止，XLSX 用 `@` 文本格式。

### UTF-8 无 BOM

中文 Windows Excel 打开无 BOM 的 UTF-8 CSV 时按 GBK 解析，中文全乱码。CSV 输出强制 `utf-8-sig`。

### 短横占位

金额列的 `-` 不是缺失值，是 0。自动检测并替换。

### 大文件性能

当前逐格 `.at[]` 循环清洗，15 万行 × 81 列需要数分钟。后续计划优化为向量化操作。

---

## 依赖

| 库 | 用途 |
|----|------|
| pandas | CSV 读取与 DataFrame 操作 |
| charset-normalizer | 首选编码检测 |
| chardet | 兜底编码检测 |
| openpyxl | XLSX 输出 |
| clevercsv（可选） | CSV 方言检测（分隔符/引号） |

```bash
pip install pandas charset-normalizer chardet openpyxl
```

---

## License

[MIT](LICENSE) © benyichan

---

**⭐ Star 这个仓库，如果它帮你从另一场 CSV 噩梦中解脱出来。**
