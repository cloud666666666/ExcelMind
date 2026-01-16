---
name: core_query
display_name: 数据查询
description: Excel 数据的基础查询、筛选、搜索、预览和统计功能
category: core
priority: 100
keywords:
  - 查询
  - 筛选
  - 过滤
  - 搜索
  - 查找
  - 找出
  - 找到
  - 显示
  - 列出
  - 看看
  - 有哪些
  - 多少
  - 预览
  - 数据
  - 记录
  - 行
  - query
  - filter
  - search
  - find
  - show
  - list
tools:
  - filter_data
  - search_data
  - get_data_preview
  - get_column_stats
  - get_unique_values
examples:
  - 帮我筛选出销售额大于1000的记录
  - 查找所有包含'北京'的数据
  - 显示前20行数据
  - 这个表有多少行
  - A列有哪些唯一值
---

# 数据查询技能

你可以使用数据查询工具来筛选、搜索和预览 Excel 数据。

## 可用工具

### filter_data
筛选数据，支持多种运算符：
- `==` 等于
- `!=` 不等于
- `>` 大于
- `<` 小于
- `>=` 大于等于
- `<=` 小于等于
- `contains` 包含
- `startswith` 以...开头
- `endswith` 以...结尾

### search_data
全文搜索，在所有列中查找包含关键词的行

### get_data_preview
预览数据的前 N 行，默认显示前 10 行

### get_column_stats
获取指定列的统计信息，包括：
- 计数、唯一值数量
- 数值列：均值、中位数、标准差、最小/最大值
- 文本列：最常见值

### get_unique_values
获取指定列的所有唯一值列表
