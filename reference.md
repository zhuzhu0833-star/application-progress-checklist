# Application Progress Checklist — Reference

## 工作表结构

| Sheet | 表名 | 行粒度 |
|-------|------|--------|
| 00 | 仪表盘 | 汇总 + 本周待办 |
| 01 | 项目总览 | 1 行 / 项目 |
| 02 | 材料进度 | 通用 5 行 + N×单项目材料 |
| 03 | 推荐信进度 | R1–R3 主数据 + 项目×推荐人 |
| 04 | Essay进度 | 1 行 / 项目 / 文书 |
| 05 | 网申账号 | 1 行 / 项目 |
| 06 | 录取进度 | 1 行 / 项目 |

## 项目 ID 规则

1. 优先使用 `config.project_id_map`（按定校 CSV「项目」列匹配）
2. 否则自动生成：`{学校缩写}-{项目缩写}`，如 `UT-MSAI`、`GT-OMSCS`
3. 冲突时追加数字后缀

### 学校缩写表

| 关键词 | 缩写 |
|--------|------|
| University of Texas / Austin | UT |
| Pennsylvania / Penn | PENN |
| Georgia Tech / Georgia Institute | GT |
| Carnegie Mellon | CMU |
| Illinois / UIUC | UIUC |
| Stanford | STAN |
| USC | USC |
| Johns Hopkins | JHU |
| Northeastern | NEU |
| Arizona State | ASU |

## 定校 CSV 字段映射

| CSV 列 | 目标 |
|--------|------|
| 截止日期 | Priority DDL |
| 学校中文名 | 学校中文名 |
| UniverisityUniversity | 学校英文名（可被 school_en_fixes 覆盖） |
| 项目 | 项目名称、ID 映射键 |
| 项目链接 | 项目链接 |
| 入学时间 | 入学批次 |
| 申请备注/录取条件 | 申请备注；含 `9/1 final` 时 Final DDL=2026-09-01 |
| 申请梯队 | 可选，预留扩展 |

## 择校 CSV enrich 字段

| CSV 列 | 写入 |
|--------|------|
| RL数量 | lor_count；`0`/`不要求` → 无 LoR |
| 申请费（美元） | fee（取末位数字） |
| 招生网址 | portal_url |
| 所属院系 | portal_system（截断 40 字） |
| 材料要求 | 推断 Essay 类型（SOP/PS/Short Answers） |

匹配逻辑：专业英文名称与定校「项目」模糊匹配，或大学英文名 + 项目名组合匹配。

## 状态枚举

**网申状态**：未注册 → 已注册 → 填写中 → 待支付 → 已提交 → Complete

**材料状态**：未开始 → 准备中 → 待审核 → 已定稿 → 已上传 → 已送分/Official → 不适用

**推荐信状态**：未邀请 → 已邀请 → 已提醒 → 已提交 → 逾期

**Essay 状态**：未开始 → 大纲中 → 写作中 → 待学生反馈 → 待顾问修改 → 已定稿 → 已上传

## 材料进度公式

主表「材料进度」= 项目ID 匹配行 + 项目ID=`通用` 行的完成比例。

完成状态：`已定稿`、`已上传`、`已送分/Official`

## 就绪度权重

| 类型 | 材料 | LoR | Essay | 网申 |
|------|------|-----|-------|------|
| default | 40% | 30% | 20% | 10% |
| no_lor | 40% | 0% | 20% | 10% |

在 config `readiness_weights` 可覆盖。

## merge 主键

`--merge` 从旧 xlsx 保留用户字段：

| Sheet | 主键 | 保留列 |
|-------|------|--------|
| 01_项目总览 | 项目ID | 网申状态、支付、提交日、卡点、负责人、更新信息 |
| 02_材料进度 | 项目ID + 材料名称 | 状态、本地路径、完成日期、备注 |
| 03_推荐信进度 | 项目ID + 推荐人编号 | 分配、邀请状态、日期、Portal 状态 |
| 03 顶部 | R1–R3 行 | 姓名、邮箱、同意状态 |
| 04_Essay进度 | 项目ID + 文书名称 | 版本、状态、字数、定稿日、链接、要点 |
| 05_网申账号 | 项目ID | 邮箱、Application ID、2FA、注册状态 |
| 06_录取进度 | 项目ID | 提交后全部用户字段 |

## config.projects 常用字段

```yaml
projects:
  UT-MSAI:
    lor_count: 3
    lor_optional_third: false
    fee: 90
    portal_system: UT Austin CDSO Apply
    portal_url: https://...
    local_dir: 申请/UT MSAI/
    essays: [SOP]
    account_note: 共用账户说明
```
