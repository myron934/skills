# Skills Repository

AI Agent 技能仓库，存放各种实用工具和自动化脚本。

## 包含的 Skills

### [pptx-to-md](./.trae/skills/pptx-to-md/)

将 PPTX 文件解析为 Markdown 格式，支持图片提取和 OCR 文字识别。

**功能特点：**
- 提取 PPT 文本内容到 Markdown
- 提取图片到本地目录（按幻灯片索引命名）
- 自动修复图片链接为相对路径
- 图片 OCR 识别（支持中英文）
- OCR 噪声自动清理

**触发条件：**
- 用户需要将 PPT/PPTX 转换为 Markdown
- 用户需要提取 PowerPoint 文件内容
- 用户需要识别 PPT 中的图片文字

## 项目结构

```
skills/
├── .trae/
│   └── skills/
│       └── pptx-to-md/       # PPTX 转 Markdown Skill
│           ├── SKILL.md      # Skill 定义文件
│           └── extract_pptx.py
├── pptx-to-md/               # 开发目录（源文件）
│   ├── SKILL.md
│   └── extract_pptx.py
├── .gitignore
├── pyproject.toml
└── README.md
```

## 安装 Skill

### 方法 1：复制到 Trae 全局 Skills 目录

```powershell
# 复制到 Trae 全局目录
Copy-Item -Recurse -Force ".trae\skills\pptx-to-md" "$env:USERPROFILE\.trae-cn\builtin\global\skills\"
```

### 方法 2：在项目中使用

将 `.trae/skills/` 目录保留在项目中，Trae 会自动识别项目级别的 Skills。

## 使用方式

安装后，直接告诉 AI 你的需求即可：

```
用户：把这个 PPT 转成 Markdown
AI：[自动调用 pptx-to-md skill]
```

或者：

```
用户：提取 presentation.pptx 的内容
AI：[自动调用 pptx-to-md skill]
```

**无需手动执行命令！**

## 开发说明

### 添加新 Skill

1. 在 `.trae/skills/` 目录创建新文件夹（如 `new-skill/`）
2. 创建 `SKILL.md` 文件，包含 frontmatter：

```markdown
---
name: "new-skill"
description: "Skill 功能描述。Invoke when 触发条件。"
---

# Skill Title

详细说明...
```

3. 添加脚本文件
4. 更新本 README

### Skill 文件格式要求

**SKILL.md 必须包含：**

| 字段 | 位置 | 说明 |
|------|------|------|
| `name` | frontmatter | Skill 唯一标识 |
| `description` | frontmatter | 功能描述 + 触发条件（<200字符） |
| 详细说明 | body | Markdown 格式的使用说明 |

## 环境配置

本项目使用 Python 虚拟环境，配置已记录在 `pyproject.toml` 中。

**虚拟环境路径：** `.venv`

**激活虚拟环境：**
```powershell
.venv\Scripts\Activate.ps1
```

## Git 忽略规则

以下内容不会被提交到 Git：
- `input/` - 输入文件目录
- `output/` - 输出文件目录
- `__pycache__/` - Python 缓存

## License

MIT
