---
name: formula
display_name: Excel 公式
description: 读取和写入 Excel 公式，支持各种 Excel 函数
category: on_demand
priority: 70
keywords:
  - 公式
  - 函数
  - SUM
  - AVERAGE
  - COUNT
  - MAX
  - MIN
  - VLOOKUP
  - IF
  - SUMIF
  - COUNTIF
  - formula
  - function
patterns:
  - "=\\w+\\("
  - "(添加|写入|设置).*(公式|函数)"
  - "(读取|查看|获取).*(公式|函数)"
tools:
  - write_formula
  - read_formula
requires:
  - core_query
examples:
  - 在 C1 添加求和公式 =SUM(A1:B1)
  - 读取 D1 单元格的公式
  - 这个单元格用的是什么公式
  - 写入 AVERAGE 函数
---

# Excel 公式技能

你可以读取和写入 Excel 公式。

## 可用工具

### write_formula
写入 Excel 公式到单元格

### read_formula
读取单元格的公式（而非计算结果）

## 支持的 Excel 函数

支持所有标准 Excel 函数，包括：
- **数学函数**: SUM, AVERAGE, COUNT, MAX, MIN, ROUND
- **逻辑函数**: IF, AND, OR, NOT
- **查找函数**: VLOOKUP, HLOOKUP, INDEX, MATCH
- **文本函数**: CONCATENATE, LEFT, RIGHT, MID, LEN
- **日期函数**: TODAY, NOW, DATE, YEAR, MONTH, DAY
- **条件函数**: SUMIF, COUNTIF, AVERAGEIF

## 注意事项

- 公式将在 Excel 中打开时计算，不会立即显示结果
- 写入公式时可以省略开头的 `=` 号，系统会自动添加
- 公式中的单元格引用区分大小写（建议使用大写）
