---
name: visualization
display_name: 数据可视化
description: 生成各类图表，包括柱状图、折线图、饼图等
category: on_demand
priority: 70
keywords:
  - 图表
  - 图
  - 柱状图
  - 折线图
  - 饼图
  - 散点图
  - 可视化
  - 画图
  - 绘制
  - 展示
  - chart
  - plot
  - graph
  - bar
  - line
  - pie
patterns:
  - "(画|绘制|生成|创建).*(图|chart)"
  - "(柱状|折线|饼|散点|雷达)图"
tools:
  - generate_chart
requires:
  - core_query
examples:
  - 画一个销售额的柱状图
  - 生成按月份的折线图
  - 用饼图展示各地区占比
---

# 数据可视化技能

你可以使用图表工具生成各类可视化图表。

## 可用工具

### generate_chart
生成图表，支持的图表类型：

| 类型 | 说明 | 适用场景 |
|------|------|----------|
| `bar` | 柱状图 | 比较不同类别的数值 |
| `line` | 折线图 | 展示趋势变化 |
| `pie` | 饼图 | 展示占比分布 |
| `scatter` | 散点图 | 展示两个变量的关系 |
| `radar` | 雷达图 | 多维度对比 |
| `funnel` | 漏斗图 | 展示转化过程 |

## 使用建议

- 柱状图适合比较离散类别的数值
- 折线图适合展示时间序列数据
- 饼图适合展示占比，类别不宜过多（建议不超过7个）
- 散点图适合展示相关性分析
