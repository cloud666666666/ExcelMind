---
name: sheet_management
display_name: 工作表管理
description: 切换、管理 Excel 工作表
category: system
priority: 60
keywords:
  - 工作表
  - sheet
  - 切换
  - 表格
patterns:
  - "切换到.+(表|sheet)"
  - "打开.+(表|sheet)"
tools:
  - switch_sheet
examples:
  - 切换到 Sheet2
  - 打开销售数据表
---

# 工作表管理技能

你可以使用工作表管理工具切换不同的工作表。

## 可用工具

### switch_sheet
切换到指定的工作表

## 使用说明

- 可以通过工作表名称切换（如 "Sheet1"、"销售数据"）
- 可以通过索引切换（从 0 开始）
- 切换后，所有数据操作将在新的工作表上进行
