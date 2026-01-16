---
name: aggregation
display_name: 聚合分析
description: 数据聚合、分组统计、排序等高级分析功能
category: on_demand
priority: 80
keywords:
  - 求和
  - 平均
  - 最大
  - 最小
  - 总计
  - 合计
  - 统计
  - 分组
  - 汇总
  - 聚合
  - 排序
  - 排名
  - sum
  - avg
  - average
  - max
  - min
  - total
  - count
  - group
  - aggregate
  - sort
  - rank
patterns:
  - "(求和|平均|最大|最小|总计|合计)"
  - "按.+分组"
  - "(升序|降序)排"
tools:
  - aggregate_data
  - group_and_aggregate
  - sort_data
requires:
  - core_query
examples:
  - 计算销售额的总和
  - 按地区分组统计销售额
  - 求出平均价格
  - 按金额降序排列
---

# 聚合分析技能

你可以使用聚合工具进行求和、平均、分组等统计分析。

## 可用工具

### aggregate_data
对单列或多列进行聚合计算，支持的聚合函数：
- `sum` 求和
- `mean` 平均值
- `count` 计数
- `min` 最小值
- `max` 最大值
- `median` 中位数
- `std` 标准差

### group_and_aggregate
按指定列分组并进行聚合统计，例如：
- 按地区分组，计算每个地区的销售额总和
- 按月份分组，计算每月的平均订单金额

### sort_data
对数据进行排序：
- 升序排列（ascending=True）
- 降序排列（ascending=False）
- 支持多列排序
