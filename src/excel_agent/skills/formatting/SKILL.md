---
name: formatting
display_name: 格式设置
description: 设置单元格格式、样式、字体、颜色、边框、合并单元格、调整行高列宽等
category: on_demand
priority: 65
keywords:
  - 格式
  - 样式
  - 字体
  - 颜色
  - 边框
  - 对齐
  - 加粗
  - 斜体
  - 下划线
  - 背景色
  - 填充
  - 合并
  - 取消合并
  - 列宽
  - 行高
  - 自动调整
  - 居中
  - 居左
  - 居右
  - format
  - style
  - font
  - color
  - border
  - merge
patterns:
  - "(设置|修改).*(格式|样式|字体|颜色|背景|边框)"
  - "把.+(加粗|变色|居中|合并)"
  - "(合并|取消合并).*(单元格|格子)"
  - "(调整|设置).*(列宽|行高)"
tools:
  - set_font
  - set_fill
  - set_alignment
  - set_border
  - set_number_format
  - set_cell_style
  - merge_cells
  - unmerge_cells
  - set_column_width
  - set_row_height
  - auto_fit_column
requires:
  - core_query
examples:
  - 把标题行加粗
  - 设置 A 列为红色
  - 给表格添加边框
  - 把 A1:C1 居中
  - 合并 A1:C1 单元格
  - 设置第一行背景为黄色
  - 调整 A 列宽度
---

# 格式设置技能

你可以设置单元格的格式和样式。

## 可用工具

### 字体设置 (set_font)
设置字体名称、大小、加粗、斜体、颜色等

### 填充设置 (set_fill)
设置单元格背景颜色

### 对齐设置 (set_alignment)
设置水平对齐、垂直对齐、自动换行

### 边框设置 (set_border)
设置边框样式、颜色

### 数字格式 (set_number_format)
设置数字显示格式：千分位、百分比、日期、货币等

### 综合样式 (set_cell_style)
一次性设置多种样式

### 合并单元格 (merge_cells / unmerge_cells)
合并或取消合并单元格

### 列宽行高 (set_column_width / set_row_height / auto_fit_column)
调整列宽、行高，或自动适应内容

## 注意事项

格式化操作需要保存文件后才能在 Excel 中看到效果。
