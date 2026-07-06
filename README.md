# application-progress-checklist

Cursor Agent Skill：从定校 CSV 生成申请进度 Checklist（Excel 7 Sheet），支持 config 配置、择校表 enrich、进度 merge 与 Google Sheets 协作。

## 安装

```bash
mkdir -p ~/.cursor/skills
git clone https://github.com/zhuzhu0833-star/application-progress-checklist.git ~/.cursor/skills/application-progress-checklist
```

或手动将本仓库复制到 `~/.cursor/skills/application-progress-checklist/`。

## 依赖

```bash
python3 -m pip install openpyxl pyyaml
```

## 使用

在 Cursor 新对话中说，例如：

- 「为 XX 同学生成申请进度 Checklist，定校表在 `选校/XX最终定校.csv`」
- 「用 application-progress-checklist 更新申请进度表并 merge 现有进度」

Agent 会读取 `SKILL.md`，创建/检查 `checklist.config.yaml`，然后运行：

```bash
python3 ~/.cursor/skills/application-progress-checklist/scripts/build_checklist.py \
  --config "/path/to/student/申请/checklist.config.yaml"

# 保留已有进度
python3 ~/.cursor/skills/application-progress-checklist/scripts/build_checklist.py \
  --config "/path/to/student/申请/checklist.config.yaml" \
  --merge "/path/to/existing.xlsx"
```

## 输入

| 文件 | 必填 | 说明 |
|------|------|------|
| `checklist.config.yaml` | 是 | 学生名、CSV 路径、项目覆盖 |
| 定校 CSV | 是 | 见 `templates/定校表_template.csv` |
| 择校 CSV | 推荐 | 自动 enrich RL 数、申请费、招生网址 |

## 输出

- `{学生名}_申请进度Checklist.xlsx` — 7 个工作表（仪表盘 / 项目 / 材料 / 推荐信 / Essay / 账号 / 录取）
- `{学生名}_申请进度Checklist_使用说明.md`

Google Sheets 协作见 [sheets-guide.md](sheets-guide.md)。

## 文件结构

```
application-progress-checklist/
├── SKILL.md
├── reference.md
├── sheets-guide.md
├── README.md
├── templates/
│   ├── config.example.yaml
│   └── 定校表_template.csv
└── scripts/
    ├── build_checklist.py
    └── merge_progress.py
```

## 同步到其他电脑

```bash
cd ~/.cursor/skills/application-progress-checklist
git pull
```

## 更新后推送

```bash
git add .
git commit -m "描述你的修改"
git push
```
