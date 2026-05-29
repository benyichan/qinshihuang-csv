<p align="center">
  <img src="assets/opening.png" alt="qinshihuang-csv" width="100%">
</p>

<h1 align="center">qinshihuang-csv</h1>

<p align="center">
  <strong>电商CSV大一统清洗方案</strong><br>
  <em>E-commerce CSV Encoding Detection & Data Cleansing Toolkit</em>
</p>

<p align="center">
  <a href="#-背景">中文</a> · <a href="#-background">English</a>
</p>

---

# 🇨🇳 背景

事情是这样的。

我之前一直在跟电商平台的CSV文件打架。淘宝导出来的GB18030，抖音导出来的UTF-8带BOM，京东给的是GBK，拼多多干脆不说自己是什么编码——你拿Excel一打开，满屏乱码。这不是最要命的。更要命的是那些**藏在字段里的脏数据**：订单号前面多了个空格、金额列里夹着￥符号、退款金额列里一个孤零零的减号" - "表示"此单无退款"——你以为是缺失值，实际上是0。

这个问题困扰了我很久。不是没有方案，是每次都得手动处理：先用记事本打开看看编码，再用公式清洗字段，最后保存的时候还得小心Excel把长数字变成科学计数法。一套操作下来，10分钟没了。如果是批量的几十个文件，半天就过去了。

所以我决定，写一个Hermes Agent的skill，把这个事情一劳永逸地解决掉。

## 踩过的坑

这个skill前后迭代了几十轮，踩的坑一个一个记录下来了。

### 坑1：编码检测不是你想的那么简单

你以为装个chardet就够了？天真。

第一版：chardet一把梭。结果碰到一个文件，chardet返回"置信度20.7%的GB18030"——低于我设的50%阈值，直接跳过了。然后回退列表一层层往下试，gb18030、gbk、gb2312全挂了——因为文件在第409字节处有一个损坏的字节。最后落到cp1252上，完美解码，出来一堆乱码。

> 起因：CSV文件在传输过程中有一个字节被写坏了。

**解决方法**：三层检测 + CJK容错回退。
1. BOM头检测（最快，0开销）
2. charset-normalizer（比chardet准）
3. chardet降门槛（检测到language='zh'时阈值从0.5降到0.2）
4. 广度回退列表（40+编码，对CJK编码用errors='replace' + 验证CJK字符数）

```python
# 核心逻辑：对CJK编码尝试容错解码
if enc in cjk_encodings:
    decoded = raw_header.decode(enc, errors="replace")
    cjk_count = sum(1 for c in decoded if '\u4e00' <= c <= '\u9fff')
    if cjk_count >= 5:
        return enc  # 至少5个汉字才接受
```

### 坑2：XLSX设了格式不等于转了类型

这个坑是最隐蔽的。

我想着，把日期列设成`YYYY-MM-DD HH:MM:SS`格式、金额列设成`#,##0.00`格式，Excel不就能正确识别了吗？

结果用户反馈：**日期筛选没有树状结构，金额左对齐**。

我查了半天，发现一个残酷的事实：**openpyxl的number_format只改变显示方式，不转换底层数据类型**。你往单元格里写一个字符串"2025-10-19 14:25:59"然后设日期格式，在Excel眼里它仍然是文本——左对齐，不可排序，筛选无树状结构。

**解决方法**：在写入XLSX之前，先把字符串转成真正的Python对象。

```python
# 日期列：先解析为datetime对象
from datetime import datetime
dt = datetime.strptime("2025-10-19 14:25:59", "%Y-%m-%d %H:%M:%S")
ws.cell(row=r, column=c).value = dt
ws.cell(row=r, column=c).number_format = 'YYYY-MM-DD HH:MM:SS'

# 金额列：先转为float
ws.cell(row=r, column=c).value = float("799.00")
ws.cell(row=r, column=c).number_format = '#,##0.00'
```

这个顺序**必须**是先转换值，再设格式——反过来设了也没用。

### 坑3：CSV长数字被科学计数法吃掉

19位的订单号`4818775826011410603`，在CSV里好好的，Excel一打开变成`4.82E+18`，后几位变成0——**数据不可逆地丢失了**。

第一版解决方案：用Excel公式写法`="4818775826011410603"`。

用户看了一眼，说："单元格显示的是=号公式，不是我要的纯数字。"

第二版：改用前导单引号`'4818775826011410603`。Excel打开时单引号自动隐藏，单元格显示纯数字，不转科学计数法。

另外，阈值从15位降到12位——因为12位的商品ID在Excel里也可能触发科学计数法。

### 坑4：CSV的UTF-8没有BOM

用户指定了utf-8编码，结果Excel打开全乱码。

原因是中文Windows版Excel打开无BOM的UTF-8 CSV时，会按本地编码（GBK）解析。**解决方案：不管用户指定utf-8还是utf-8-sig，CSV输出统一用utf-8-sig**。

### 坑5：金额列里的减号不是减号

电商CSV里，退款金额列常常出现一个孤零零的`-`——这不是数据缺失，这是"此单无退款"，值是0。

所以我加了一个逻辑：在数值格式列中检测纯减号/短横/破折号，替换为"0"。

```python
if col in numeric_set and is_pure_dash(cleaned):
    df_clean.at[idx, col] = "0"  # 替换为字符串"0"
```

## 架构

最终沉淀下来的架构是这样的：

### 四层编码检测管线

```
BOM头检测 → charset-normalizer → chardet → 40+编码回退列表
```

### 三步清洗流程

```
① 自动识别表头（扫描前5行）→
② 扫描前15行数据，逐列确定清洗策略和格式类型 →
③ 全量清洗 + 值类型转换 + 输出
```

### 列格式优先级

```
文本格式（@，12+位纯数字）> 日期格式（YYYY-MM-DD HH:MM:SS）> 数值格式（#,##0.00）
```

互斥，避免同一列被标记多个格式。

## 环境

| 依赖 | 用途 |
|------|------|
| Python 3.13+ | 运行环境 |
| pandas | CSV读写、DataFrame处理 |
| charset-normalizer | 主编码检测 |
| chardet | 编码检测兜底 |
| openpyxl | XLSX写入（Excel 2016+兼容） |
| clevercsv（可选） | CSV方言检测（分隔符/引号） |
| python-docx | 文档读取（仅测试用） |

## 反馈 / 已知问题

- **大文件性能**：当前逐格循环清洗（.at[]）在10万行×80列以上会显著变慢，后续考虑vectorized方案
- **Chardet中文置信度**：GB18030编码的文件被chardet检测时置信度可能低至0.2，已在代码中特殊处理
- **CleverCSV网络安装**：Windows下pip install clevercsv可能因SSL证书问题失败，不影响核心功能

---

# 🇬🇧 Background

The story goes like this.

For a long time, I've been fighting with CSV files exported from Chinese e-commerce platforms. Taobao exports GB18030, Douyin prefers UTF-8 with BOM, JD.com uses GBK, and Pinduoduo doesn't bother telling you what encoding it uses — open any of these in Excel and you get a screen full of garbled text. That's not even the worst part. The real nightmare is the **hidden dirty data** lurking inside the fields: extra spaces before order IDs, ¥ symbols embedded in price columns, a lonely dash `-` in refund amounts that actually means "zero refund" — not a missing value, just zero.

This problem haunted me for a long time. Not because there's no solution, but because every time I had to manually go through the same tedious process: open the file in Notepad++ to figure out the encoding, use formulas to clean up the fields, and pray that Excel wouldn't turn my 19-digit order IDs into scientific notation. A single file takes 10 minutes. A batch of dozens? That's half a day gone.

So I decided to write a Hermes Agent skill and solve this problem once and for all.

## The Pitfalls

This skill went through dozens of iterations. Here are the battles I fought.

### Pitfall 1: Encoding Detection Is Not What You Think

First version: just use chardet. Naive.

I hit a file where chardet returned "GB18030 with 20.7% confidence" — below my 50% threshold, so it was skipped. The fallback list tried gb18030, gbk, gb2312 — all failed, because there was **one corrupted byte** at position 409. Finally it fell through to cp1252, which decodes everything without throwing errors, and produced beautiful gibberish.

**Fix**: 4-layer detection with CJK-tolerant fallback.

### Pitfall 2: XLSX Format != Type Conversion

I thought setting `number_format = 'YYYY-MM-DD HH:MM:SS'` on date columns and `#,##0.00` on amount columns would be enough.

The user's feedback: "Date filters don't show hierarchy. Amounts are left-aligned."

I spent hours debugging before realizing the brutal truth: **openpyxl's `number_format` only changes display, not the underlying data type**. Writing a string `"2025-10-19"` into a cell and setting date format makes it a text cell in Excel's eyes — left-aligned, unsortable, unfilterable.

**Fix**: Convert strings to proper Python objects BEFORE writing to XLSX.

### Pitfall 3: Scientific Notation Eats Long Numbers

A 19-digit order ID like `4818775826011410603` looks perfect in the CSV file. But when Excel opens it, it becomes `4.82E+18`, with the last digits turned to zeros — **data loss is permanent**.

First fix: Excel formula notation `="4818775826011410603"`.

User: "The cell shows the formula, not the number I want."

Second fix: Leading single quote `'4818775826011410603`. Excel hides the quote, shows the pure number, no scientific notation.

Also lowered the threshold from 15 to 12 digits — 12-digit product IDs can trigger scientific notation too.

### Pitfall 4: UTF-8 Without BOM Breaks Excel

User specified UTF-8 encoding. Excel opened it as pure garbage.

The reason: Chinese Windows Excel interprets BOM-less UTF-8 as the local encoding (GBK). **Fix: Always use utf-8-sig for CSV output, regardless of what the user specified.**

### Pitfall 5: That Dash in the Amount Column Isn't a Dash

In e-commerce CSVs, the refund amount column often contains a single `-` character. This isn't missing data — it means "no refund for this order", value = 0.

Added a simple rule: detect pure dash/hyphen characters in numeric columns and replace with "0".

## Architecture

### 4-Layer Encoding Detection Pipeline

```
BOM Detection → charset-normalizer → chardet → 40+ Encoding Fallback List
```

### 3-Step Cleaning Flow

```
① Auto-detect header (scan first 5 rows) →
② Scan first 15 data rows per column →
③ Full clean + value type conversion + output
```

### Column Format Priority

```
Text Format (@, 12+ digits) > Date Format (YYYY-MM-DD HH:MM:SS) > Numeric Format (#,##0.00)
```

Mutually exclusive — no column gets tagged with multiple formats.

## Environment

| Dependency | Purpose |
|------------|---------|
| Python 3.13+ | Runtime |
| pandas | CSV I/O, DataFrame |
| charset-normalizer | Primary encoding detection |
| chardet | Encoding detection fallback |
| openpyxl | XLSX output (Excel 2016+) |
| clevercsv (optional) | CSV dialect detection |
| python-docx | Doc reading (test only) |

## Known Issues

- **Large file performance**: Cell-by-cell `.at[]` iteration slows down significantly above 100K rows × 80 cols. Vectorized approach planned.
- **Chardet CJK confidence**: GB18030 files can get as low as 0.2 from chardet, handled specially in code.
- **CleverCSV installation**: May fail under Windows due to SSL certificate issues — core functionality unaffected.

---

**⭐ If this project saved you time, give it a star!**

> Inspired by years of wrestling with e-commerce CSV exports. Built with ❤️ for every data analyst who's ever opened a CSV file and screamed.
