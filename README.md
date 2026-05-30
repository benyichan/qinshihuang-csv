<p align="center">
  <img src="assets/opening-zh.png" alt="qinshihuang-csv" width="100%">
</p>

<h1 align="center">qinshihuang-csv</h1>

<p align="center">
  <strong>电商 CSV 编码检测与数据清洗工具</strong><br>
  <em>E-commerce CSV encoding detection & data cleansing toolkit</em>
</p>

<p align="center">
  <a href="README_EN.md">English</a> · <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="MIT"></a>
</p>

---

## 这是什么

国内电商平台导出的 CSV 文件普遍存在两个问题：

1. **编码不统一**。同一个文件用 Excel 打开是乱码，用 WPS 打开正常。不是文件坏了，是解析规则不匹配。
2. **字段里有脏数据**。肉眼看不见的：首尾空格、币种符号嵌入金额、长数字被 Excel 转科学计数法导致精度丢失、退款列用单个 `-` 表示零值。

这个 skill 就是为了自动解决这些问题——把不同编码的 CSV 文件统一清洗成标准格式。

## 技术栈

| 依赖 | 用途 |
|------|------|
| Python 3.13+ | 运行环境 |
| pandas | CSV 读写与 DataFrame 处理 |
| charset-normalizer | 主要编码检测引擎 |
| chardet | 编码检测兜底 |
| openpyxl | XLSX 输出（Excel 2016+ 兼容） |
| clevercsv（可选） | CSV 方言检测（自动识别分隔符/引号） |

## 快速开始

```bash
pip install pandas charset-normalizer chardet openpyxl

# 处理单个文件
python scripts/qinshihuang_cleaner.py 你的文件.csv --to-xlsx

# 批量处理目录
python scripts/qinshihuang_cleaner.py 目录路径/ --batch --to-xlsx
```

Hermes Agent 用户直接加载 skill 后告诉 agent 即可。

## 核心能力

### 四层编码检测

```
BOM 头检测 → charset-normalizer → chardet（中文降阈至0.2） → 40+ 编码回退列表
```

### 三步清洗

```
① 自动定位表头 → ② 逐列分析脏数据类型 → ③ 全量清洗 + 输出
```

### 列格式自动识别（互斥优先级）

```
文本格式（@，≥12位纯数字） > 日期格式（YYYY-MM-DD HH:MM:SS） > 数值格式（#,##0.00）
```

## 踩过的坑

### 编码检测

chardet 单独不够用。实测中遇到过中文编码文件置信度仅 20.7% 的情况，且文件内含一个损坏字节导致所有 GBK/GB18030 解码失败，最后回退到 cp1252 输出乱码。最终方案是四层管线 + CJK 容错回退：对中文编码先用 `errors='replace'` 解码，再验证汉字数量来决定是否接受。

### XLSX 格式不转类型

openpyxl 的 `number_format` 只改变显示方式，不改变底层数据类型。往单元格写字符串 `"2025-10-19"` 再设日期格式，Excel 仍然按文本处理（左对齐、不可排序）。必须先转换值类型（`datetime.strptime` → datetime 对象，`float()` → 数值），再设格式。

### 长数字精度丢失

超过 12 位的纯数字被 Excel 转科学计数法后精度不可逆丢失。CSV 输出用前导单引号 `'123456789012345` 保护，Excel 打开时单引号自动隐藏。阈值设为 12 位而非 15 位，因为商品 ID 等 12 位数字也可能被 Excel 误转。

### UTF-8 无 BOM

中文 Windows 下 Excel 打开无 BOM 的 UTF-8 CSV 会按本地编码（GBK）解析产生乱码。无论用户指定什么编码，CSV 输出统一用 `utf-8-sig`。

### 短横占位

金额列中常见单个 `-` 符号，不是缺失值而是零值。在数值格式列中自动检测并替换为 `"0"`。

## 已知问题

- **大文件性能**：逐单元格清洗在 10 万行以上明显变慢，后续考虑向量化方案
- **chardet 中文置信度**：中文编码文件被 chardet 检测时置信度可能低至 0.2，已通过在代码中降阈处理
- **Windows 安装**：部分依赖（clevercsv）可能因 SSL 证书问题安装失败，不影响核心功能

---

**⭐ 如果这对你有帮助，给个星吧。**

[MIT](LICENSE) © benyichan
