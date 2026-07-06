---
name: application-progress-checklist
description: >-
  从定校 CSV 生成申请进度 Checklist（Excel 7 Sheet：仪表盘/项目/材料/推荐信/Essay/账号/录取）。
  支持 config.yaml 补充 RL 数与 Essay 类型、择校表 enrich、进度 merge、Google Sheets 协作说明。
  Use when user asks for 申请进度表, 申请 checklist, 定校进度跟踪, application progress tracker,
  生成申请协作表, 申请进度 skill.
---

# Application Progress Checklist

从定校 CSV 为签约学生生成可协作的申请进度 Excel，结构与 [reference.md](reference.md) 一致。

## Quick start

```bash
python3 -m pip install openpyxl pyyaml

python3 ~/.cursor/skills/application-progress-checklist/scripts/build_checklist.py \
  --config "/path/to/student/申请/checklist.config.yaml"

# 保留已有进度（定校表更新后）
python3 ~/.cursor/skills/application-progress-checklist/scripts/build_checklist.py \
  --config "/path/to/student/申请/checklist.config.yaml" \
  --merge "/path/to/student/申请/现有Checklist.xlsx"
```

## Workflow checklist

```
- [ ] 确认定校 CSV 路径与项目行数
- [ ] 复制 templates/config.example.yaml → 学生 申请/checklist.config.yaml
- [ ] 填写 student_name、school_list_csv；可选 research_csv
- [ ] 设置 project_id_map（保持稳定 ID）与 projects.* 覆盖
- [ ] 运行 build_checklist.py（有旧表则加 --merge）
- [ ] 验证：项目数、通用材料 5 行、无 LoR 项目显示 N/A、待办公式无 #ERROR
- [ ] 交付 xlsx + 使用说明；提醒 Sheets 导入后见 sheets-guide.md 自检
```

## 输入文件

| 文件 | 必填 | 说明 |
|------|------|------|
| `checklist.config.yaml` | 是 | 学生名、路径、项目覆盖 |
| 定校 CSV | 是 | 见 templates/定校表_template.csv |
| 择校 CSV | 推荐 | 自动 enrich RL 数、申请费、招生网址 |

定校 CSV 最小列：`截止日期, 学校中文名, UniverisityUniversity, 项目, 项目链接, 入学时间, 申请备注/录取条件`

## 输出

- `{学生名}_申请进度Checklist.xlsx` — 7 个工作表
- `{学生名}_申请进度Checklist_使用说明.md` — 简要协作说明

## 设计要点

- **材料合并**：项目ID=`通用` 的 5 项（成绩单/TOEFL/Resume）所有学校共用
- **单项目材料**：网申表填写、Background 说明、申请费（每校一行）
- **Short Answers**：仅在 Essay 表跟踪，不在材料表重复
- **无 LoR 项目**：`lor_count: 0` → 推荐信进度显示 N/A，就绪度权重自动切换

## Google Sheets

协作上传与公式自检见 [sheets-guide.md](sheets-guide.md)。

## 新增学生步骤

1. 在学生 `申请/` 目录创建 `checklist.config.yaml`（从 templates 复制）
2. 指向该生定校 CSV / 择校 CSV
3. 在 `projects:` 下配置 `local_dir`、`essays`、`lor_count` 等
4. 运行脚本；首次无需 `--merge`

## 定校表变更

有进度数据时 **必须** 使用 `--merge`；merge 按主键保留用户字段，详见 [reference.md](reference.md#merge-主键).

## 依赖

```bash
python3 -m pip install openpyxl pyyaml
```

## 附加资源

- 列定义与状态枚举：[reference.md](reference.md)
- Google Sheets 协作：[sheets-guide.md](sheets-guide.md)
- 配置模板：[templates/config.example.yaml](templates/config.example.yaml)
