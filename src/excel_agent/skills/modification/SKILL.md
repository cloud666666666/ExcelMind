---
name: modification
display_name: 数据修改
description: 写入、修改、删除 Excel 数据，包括单元格写入、批量写入、行列操作、保存文件
category: on_demand
priority: 75
keywords:
  - 写入
  - 修改
  - 更新
  - 删除
  - 添加
  - 插入
  - 加上
  - 保存
  - 另存
  - 导出
  - 覆盖
  - 原始文件
  - 末尾
  - 结尾
  - 最后
  - 新增
  - 追加
  - write
  - update
  - delete
  - insert
  - save
  - export
  - append
patterns:
  - "(写入|修改|更新|删除).+"
  - "把.+(改成|设为|设置为)"
  - "在.+(添加|插入|加上)"
  - "(末尾|结尾|最后).*(加上|添加|写入)"
  - "(加上|添加).*(合计|总计|汇总)"
  - "保存(文件|表格|到原始)?"
  - "导出(到|为)?"
tools:
  - write_cell
  - write_range
  - insert_rows
  - delete_rows
  - save_file
  - save_to_original
  - export_file
  - quick_export
  - get_change_log
requires:
  - core_query
examples:
  - 把 A1 单元格写入 100
  - 在 A1 开始写入数据
  - 删除第 5 行
  - 插入 3 行
  - 保存文件
  - 保存到原始文件
  - 导出到新文件
---

# 数据修改技能

你可以使用数据修改工具来写入和修改 Excel 数据。

## 可用工具

### write_cell
写入单个单元格的值

### write_range
批量写入数据到指定范围

### insert_rows
在指定位置插入空白行

### delete_rows
删除指定行

### save_file
保存文件到工作副本

### save_to_original
覆盖保存到原始文件（慎用）

### quick_export
快速导出到原文件所在目录（推荐）

### export_file
导出到指定位置

### get_change_log
查看所有修改记录

## 重要说明

- 所有修改默认保存到工作副本，不会影响原始文件
- 使用 `save_file` 保存到副本
- 使用 `save_to_original` 覆盖原始文件（慎用！）
- 使用 `quick_export` 快速导出到原文件所在目录（推荐）
- 使用 `export_file` 导出到指定位置
- 可以使用 `get_change_log` 查看所有修改记录
