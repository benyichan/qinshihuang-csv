---
name: qinshihuang-csv
category: data-science
description: 电商CSV文件全流程清洗——编码检测/转换 + 字段级脏数据清洗。自动识别编码，清除首尾特殊字符，防止科学计数法，支持输出CSV或XLSX(Excel 2016+)
trigger: csv|csv文件|清洗|csv清洗|编码转换|convert_csv|qinshihuang
---

# qinshihuang-csv — 电商CSV大一统清洗方案

## 适用场景

用户拿到一个或多个电商平台导出的 CSV 文件，需要：
1. **识别编码**并统一转换为目标编码
2. **清洗字段脏数据**（首尾特殊字符、不可见字符、科学计数法防护）
3. 输出为 `{原名}_cleaned.csv` 或 `{原名}_cleaned.xlsx`

## 执行前需要向用户确认的问题

**⚠️ 这是强制规则，每次执行前必须提问，不得跳过。**

执行脚本前，先依次问用户这 4 个问题，**每个问题独立提问，不要拼成组合选项**：

1. **目标编码？** 默认 `utf-8-sig`（带BOM，兼容Excel）。用户可自由指定任意编码（utf-8、gbk、gb18030 等）
2. **是否转存 .xlsx？** 如果否，只输出 `_cleaned.csv`；如果是，额外输出 `_cleaned.xlsx`
3. **输出文件存放路径？** 默认当前用户的桌面，让用户确认或指定
4. **是否需要处理整个目录？** 单个文件还是批量处理某目录下所有 `.csv`

**例外情况：** 如果用户明确说"直接执行"、"操练起来"、"按默认来"等，可用上次用户确认过的设置直接执行。
但如果用户只给出了文件路径、没有回答过上述问题，必须先问。

## 执行纪律（强制遵守）

1. **遵循 skill 标准流程** — 严格按照 SKILL.md 定义的流程执行，包括先提问再执行。
2. **禁止擅自跳过标准操作** — 如果按 skill 标准操作无法达到目的，必须先排查原因、尝试在 skill 框架内修复。
   严禁不经过排查就跳过标准操作，转而使用 skill 没有写明的替代方案。
3. **替代方案需用户同意** — 如果经过排查确认 skill 框架无法满足需求，需要采用替代方案，
   必须向用户说明原因并获得同意后方可执行。
4. **发现 bug 立即修复** — 执行中发现 skill 本身的 bug（代码逻辑错误、参数错误、遗漏场景等），
   应在排查确认后立即修复，沉淀到 skill 中，避免下次再踩。

## 执行流程

### Step 1: 编码检测与转换

```
输入文件(.csv)
  ↓ ① BOM头检测 (检查前4字节)
  ↓ ② charset-normalizer 识别 (sample_size=100000, confidence≥0.7)
  ↓ ③ chardet 兜底 (与 step② 交叉验证)
  ↓ ④ 广度回退列表尝试 (40+编码)
  ↓ ⑤ 如果已目标编码 → 询问用户是否继续
  ↓ ⑥ 解码成功 → 加载为 DataFrame (所有列读为 str)
  ↓ (内存中的 DataFrame 进入 Step 2)
```

**回退编码列表**（按优先级，去重后依次尝试）：
`utf-8-sig`, `utf-8`, `gb18030`, `gbk`, `gb2312`, `big5`, `big5hkscs`, `shift_jis`, `euc-jp`, `iso-2022-jp`, `euc-kr`, `utf-16`, `utf-16-le`, `utf-16-be`, `utf-32`, `cp1252`, `iso-8859-1`, `iso-8859-15`, `latin1`, `mac_roman`, `cp437`, `cp850`, `koi8-r`, `iso-2022-kr`, `hz`, `iso-2022-cn`, `euc-tw`, `gb18030-2022`

> 注意：`charset-normalizer` 检测到 GB2312 时统一映射为 GB18030（超集兼容）

### Step 2: 字段级数据清洗

```
DataFrame (所有列为 str 类型)
  ↓ ① 自动识别表头 (规则：从第0行扫描，找第一个不含脏数据的纯文本行即为表头)
  ↓ ② 取前15行数据，逐列扫描脏数据特征 + 列类型推断：
  ↓    - 哪些列含 "首尾空格/逗号/分号/制表符/币种符号" → 需要清洗
  ↓    - 哪些列清洗后是 "12+位连续纯数字" → 文本格式列（防科学计数法）
  ↓    - 哪些列匹配日期模式 (YYYY-MM-DD HH:MM:SS 等5种) → 日期格式列
  ↓    - 哪些列含小数点+可解析为数值 → 数值格式列（金额）
  ↓    - 哪些列有纯短横/减号 `-` → 标记为短横替换列（替换为0）
  ↓ ③ 列格式优先级：文本 > 日期 > 数值（互斥，避免同一列标记多个格式）
  ↓ ④ 全量清洗（逐列应用策略）
  ↓ ⑤ 短横列中的纯 `-`/`–`/`—` 替换为 "0"
  ↓ ⑥ 输出文件
```

**列格式优先级**：
1. **文本格式** (`@`) — 防科学计数法，最高优先级。12+位纯数字列自动标记
2. **日期格式** (`YYYY-MM-DD HH:MM:SS`) — 排除已标记为文本的列
3. **数值格式** (`#,##0.00`) — 仅含小数点的列可被标记，排除文本和日期列

**清洗规则详情**（按应用顺序）：
1. 移除单元格内的制表符 `\t`、换行符 `\n`、回车符 `\r`
2. **循环清除首尾特殊字符**（因为可能连续出现，如 `,,，空格￥1234567890123456789￥，空格`）：
   - 英文逗号 `,`
   - 英文分号 `;`
   - 空格（一个或多个）
   - 币种符号：`￥` `$` `¥` `€` `£` `₩` `₽` `₪` `₫` `₱` `₹` `₴` `₸` `₺` `₼` `₿`
   - 不可见控制字符（Unicode category Cc 和 Cf）
3. 清洗后如果列被标记为"文本格式列"→记录为后续写入时的格式要求

### Step 3: 输出

- CSV 输出：`{原名}_cleaned.csv`
  - **编码**：如果用户指定 `utf-8`，CSV 文件自动升级为 `utf-8-sig`（带 BOM）。
    这是必需的——Excel（特别是中文 Windows 版）打开无 BOM 的 UTF-8 CSV 时，
    会按系统本地编码（GBK）解析，中文全变乱码。
  - **文本格式列**（12+位纯数字）使用**前导单引号** `'123456789012345` 写法。
    Excel 打开时单引号不显示、单元格保持文本、不转科学计数法。
    千万不要用 `="value"` 公式写法（用户明确反对，因为公式显示奇怪）。
- XLSX 输出：`{原名}_cleaned.xlsx`，用 openpyxl，Excel 2016+ 兼容
  - **关键：必须先转换值类型，再设格式。** 只设 `number_format` 而不转换值类型（datetime/float），
    Excel 仍会按文本处理（左对齐、筛选无树状结构）。
  - **文本格式** (`@`) — 值写入前加前导 `'`，再设 `@` 格式
  - **日期格式** (`YYYY-MM-DD HH:MM:SS`) — 字符串先解析为 `datetime.datetime` 对象，再设日期格式
  - **数值格式** (`#,##0.00`) — 字符串先转为 `float`，再设数值格式
  - 列宽自适应（表头+前100行最大值，上限50字符）
- 让用户确认或输入保存路径（默认桌面）

## 实现架构细节

实现笔记和决策记录见 `references/implementation-notes.md`，输出格式决策见 `references/output-format-decision-log.md`。
项目三层架构说明（Skill / Prompt / 源码+exe）见 `references/project-three-layer-architecture.md`。

## 项目位置

`D:\\projects\\qinshihuang-csv\\` — 已开源至 GitHub `benyichan/qinshihuang-csv`。

## 三层递进交付结构

项目按目标用户分为三个递进层次，详见 `references/project-architecture.md`：

```
Layer 1 — Skill（核心 CLI，Hermes Agent 调用）
  scripts/qinshihuang_cleaner.py

Layer 2 — Prompt（教程，教零基础用户在 Trae 中搭建环境）
  prompts/trae-python-setup-prompt.txt

Layer 3 — 源码 + 二进制（即开即用）
  src/秦始皇CSV清洗器_源码.py  ← 完整 tkinter GUI 桌面应用
  dist/秦始皇CSV清洗器.exe     ← PyInstaller 打包，双击运行
```

## 入口选择

| 场景 | 入口 | 说明 |
|------|------|------|
| Hermes Agent 内部调用 | `scripts/qinshihuang_cleaner.py` | CLI，有 log 输出，适合 agent 接管 |
| 用户直接运行源码 | `python src/秦始皇CSV清洗器_源码.py` | 需 Python + 依赖 |
| 用户双击运行 exe | `dist/秦始皇CSV清洗器.exe` | 无 Python 环境也可用 |

## 依赖

```bash
pip install charset-normalizer chardet pandas openpyxl
```

（这些依赖已经在环境里装好了）

## 可选依赖

```bash
pip install clevercsv
```
如果安装了 CleverCSV，脚本会自动用它做精确的方言检测（分隔符、引号字符、转义字符）。不安装也不影响——会回退到简单启发式检测。

## 参考的 GitHub 仓库

- **CleverCSV** (alan-turing-institute/CleverCSV) — 方言检测（分隔符/引号/转义字符）+ 编码检测。核心 `detect_dialect()` 用于自动识别 CSV 格式参数，`detect_encoding()` 封装了 chardet。
- **chardet** (chardet/chardet) — 经典编码检测库，作为第三层兜底。
- **csv-detective** (datagouv/csv-detective) — 列语义类型推断（清洗后校验用，非核心流程）。
- **csvprofiler** (LarryKuhn/CSV-Profiler) — CSV 统计分析/HTML 报告（锦上添花，非必要）。

## 踩坑记录 / 注意事项

1. **GB2312 → GB18030 映射**：charset-normalizer 和 chardet 都可能把 GBK 文件识别为 GB2312。由于 GB18030 完全兼容 GBK 和 GB2312，检测到 gb2312/gbk 时统一当 gb18030 处理。
2. **BOM 头**：utf-8-sig 读取时会自动跳过 BOM，写入时自动添加 BOM。这对 Excel 兼容性至关重要。
3. **科学计数法预扫描**：必须先扫描判断哪些列是"特殊字符+12+数字"，再清洗、再写入。顺序错会导致：
   - 如果先清洗再判断：特殊字符被移除后数字暴露出来→才触发保护→但 CSV/XLSX 已用文本写入，无效
   - 如果设格式但不转值类型：XLSX 中 `@`/`#,##0.00`/`YYYY-MM-DD HH:MM:SS` 仅改变显示方式，
     不转换底层值类型（datetime/float），Excel 仍按文本处理（左对齐、筛选无树状结构）
   - **正确顺序**：扫描→标记列类型→清洗→写入时按列类型转换值+设格式
4. **pandas read_csv 的 dtype=str**：必须强制所有列为 str 类型，防止长数字被 pandas 自动转成 float/int。
   同时用 `keep_default_na=False, na_filter=False` 防止空单元格被转成 NaN。
5. **循环清除**：特殊字符可能复合出现（如 `  ,,￥  123`），单次 strip 不够——清除首尾后可能暴露新的特殊字符位置。
   必须用 `while prev != text` 循环直到无变化。
6. **XLSX 值类型转换**：openpyxl 的 `number_format` 只改变显示格式，不转换底层值类型。
   字符串 `"2025-10-19"` 设 `YYYY-MM-DD` 格式 → 仍是文本（左对齐，无法排序筛选）。
   必须先用 `datetime.datetime.strptime()` 解析为 datetime 对象，再写入 cell。
   同理金额列必须 `float(val)` → 数值后设 `#,##0.00` → 右对齐可计算。
7. **on_bad_lines='skip'**：真实电商CSV常有字段内未转义的逗号导致列数不匹配（如尾随逗号、地址字段含逗号）。
   `on_bad_lines='skip'` 跳过残损行，避免整个文件读不进来。用 `header=None` 避免 pandas
   把错误行数判定为解析失败。
8. **header=None 读取**：不要用 pandas 默认的 `header=0`。电商CSV的表头可能有前置空行/备注行。
   用 `header=None` 读所有行，再用 `auto_detect_header()` 扫描前5行定位真实表头。
9. **分隔符检测**：淘宝系CSV常用制表符 `\t` 分隔，其他平台用逗号 `,`。
   如果有 CleverCSV，优先用 `detect_dialect()` 做精确检测（delimiter/quotechar/escapechar）。
   否则读第一行检查是否有 `\t` 且无 `,` 来判断。
10. **XLSX 空单元格**：openpyxl 写入时空字符串单元格变成 None，不影响格式但看起来和原始不同。
    这在 Excel 中打开时无实际影响。
11. **chardet 对中文置信度偏低**：中文CSV文件（特别是含坏字节的）被 chardet 检测时置信度可能低至
    0.2-0.3，但其 `language='zh'` 字段依然正确。检测时对 `language='zh'` 降门槛到 0.2。
12. **charset-normalizer API 版本差异**：不同版本对 `from_path()` 的参数要求不同（有的版本不支持
    `sample_size` 参数）。需同时备有 `from_bytes(raw_sample).best()` 回退路径。
13. **坏字节导致 CJK 编码被跳过**：GB18030/GBK 文件的少量损坏字节会使严格解码失败，
    回退列表继续遍历直到 cp1252/latin1 等"什么字节都能解码"的编码匹配成功，输出中文乱码。
    对策：
    - 回退列表：对 CJK 编码（gb18030/gbk/big5等）用 `errors='replace'` + 验证 CJK 字符数
      （≥5个汉字才接受）
    - chardet 检测：`language='zh'` 降门槛到 0.2
    - pandas 读取：`encoding_errors` 从 `'strict'` 自动回退到 `'replace'`
14. **金额列中的纯短横占位**：电商CSV中金额列常见纯 `-` 或 `–` 表示"此订单无退/无佣金"。
    不是数据缺失，而是值=0。清洗时必须替换为 `"0"`。
15. **CSV 文本格式必须用前导单引号，不能用公式**：第一次实现用了 `="123..."`，结果 Excel 单元格
    显示的是 `="123..."` 公式本身而不是数字。用户明确反对。改为前导单引号 `'123...`，
    Excel 把 `'` 识别为文本前缀，显示时不显示引号，单元格内容保持纯字符串。
    注意：`pandas.to_csv()` 不会额外转义前导单引号，所以 CSV 中写为 `,'123...,` 即可。
16. **12位数字保护阈值**：Excel 只能精确表示 15 位有效数字，但订单号/流水号等常见 12-19 位数字
    都应保护。实际验证：12 位的商品ID（如 `955956980424`）在 Excel 中已可能触发科学计数法。
    阈值设为 12 位（含）。
17. **`--batch` 参数的效果**：不带 `--batch` 时输出说明文字提示用户，仍执行批量处理。
    带 `--batch` 时静默执行。两种情况的处理逻辑相同。
18. **UTF-8 无 BOM → Excel 中文乱码**：用户指定 `utf-8` 时，CSV 输出必须自动升级为 `utf-8-sig`（带 BOM）。
    中文 Windows 的 Excel 打开无 BOM 的 UTF-8 CSV 时，会按系统本地编码（GBK）解析，中文字符全部乱码。
    这不是用户选择问题，是 Excel 的实现缺陷，必须无条件修复。
19. **大文件性能瓶颈**：当前 `clean_dataframe()` 使用双层 `.at[]` 循环逐格清洗，
    对小文件（<1 万行）无影响，但 15 万行 × 81 列（~1200 万单元格）需要数分钟且中间无进度日志。
    如果后续需要处理大文件，应优化为向量化操作（`.apply()` 或字符串操作），
    或新增分块处理模式（chunking），并添加中间进度日志。