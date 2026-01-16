---
name: calculation
display_name: 数学计算
description: 执行数学计算和表达式求值
category: on_demand
priority: 50
keywords:
  - 计算
  - 算
  - 加
  - 减
  - 乘
  - 除
  - 等于
  - 表达式
  - 公式
  - 数学
  - calculate
  - compute
  - math
patterns:
  - "\\d+\\s*[\\+\\-\\*\\/]\\s*\\d+"
  - "计算.+"
tools:
  - calculate
examples:
  - 计算 100 * 1.5 + 200
  - 1000 / 4 等于多少
---

# 数学计算技能

你可以使用计算工具执行数学运算。

## 可用工具

### calculate
执行数学表达式计算

## 支持的运算

### 基本运算
- `+` 加法
- `-` 减法
- `*` 乘法
- `/` 除法

### 内置函数
- `abs()` 绝对值
- `round()` 四舍五入
- `min()` 最小值
- `max()` 最大值
- `sum()` 求和
- `pow()` 幂运算

## 使用示例

```
计算 100 * 1.5 + 200
=> 350.0

计算 round(123.456, 2)
=> 123.46

计算 max(10, 20, 30)
=> 30
```
