# 企业实习报告撰写 Skill（internship-report-writing）

面向**电子科技大学（UESTC）学生**的企业实习**中期报告与终期总结报告**撰写技能。支持多项目、多职责场景，覆盖信息采集、章节规划、去 AI 化写作、参考文献顺序管理，并能**基于学校下发的 Word 模板或已确认终期范例一键生成排版好的 `.docx` 文档**。

适用于 Cursor、Claude Code 等支持 Agent Skills 的环境。

## 它能做什么

- 按学校模板结构组织章节，不自行增删
- 终期总结报告采用九章结构，建立“初期任务 - 中期状态 - 终期结果”的任务闭环
- 多项目并行：分别梳理每个项目本人负责的模块与贡献
- 读取本地代码仓库，把口述工作还原成真实技术细节，避免笼统失真
- 去 AI 化：输出客观、严谨的学术风格文本
- 参考文献按正文首次出现顺序编号，文末列表一一对应
- 自动把 Mermaid 图渲染为图片并嵌入文档
- 严格套用模板的标题 / 正文 / 参考文献字体字号，直接交付 `.docx`

## 目录结构

```
internship-report-writing/
├── SKILL.md              # 技能主文件（工作流、写作规范、格式规格）
├── template.md           # 中期报告各章节撰写要点说明
├── template-final.md     # 终期总结报告各章节撰写要点说明
├── references/
│   └── final-report-guide.md # 终期范例的结构、版式与终检规则
├── scripts/
│   └── build_docx.py     # Markdown 章节 + Word 模板 → 排版好的 .docx
└── README.md
```

## 安装

### Cursor

将本仓库克隆到项目的 `skills/` 目录，或用户级技能目录：

```bash
git clone https://github.com/iJcher/UESTC-intership_report_writing-skill.git skills/internship-report-writing
```

### Claude Code

克隆到 `.claude/skills/` 下即可被自动发现：

```bash
git clone https://github.com/iJcher/UESTC-intership_report_writing-skill.git .claude/skills/internship-report-writing
```

## 使用

1. 向 AI 说明要写中期报告或终期总结报告，并按 `SKILL.md` 顶部「需要用户提供的信息」清单提供：报告类型、学校模板/终期范例 `.docx` 路径、公司/部门/岗位、项目清单，**以及最关键的各项目本地代码仓库路径**。终期报告还需提供初期和中期报告（或其中的任务、进度和后续计划摘录）。
2. AI 按工作流采集信息、规划章节、逐章撰写并整理参考文献。
3. 生成 `.docx`：

```bash
pip install python-docx
python scripts/build_docx.py \
  --template "学校模板.docx" \
  --input chapters \
  --output "企业实习报告.docx" \
  --title "企业实习中期报告"
```

> 生成图表需联网（脚本调用 Kroki / mermaid.ink 将 Mermaid 渲染为图片）。

## 环境要求

- Python 3.8+
- `python-docx`
- 联网（仅图表渲染阶段需要）
