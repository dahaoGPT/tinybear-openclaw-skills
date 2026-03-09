# 将 Claude Code pptxmaker Skill 改造为 Antigravity Skill

将 `.agent/skills/pptxmaker/` 下的 Claude Code 风格的 skill 改造为符合 Antigravity 规范的 skill 格式。

## 两种格式的关键差异

| 特性 | Claude Code | Antigravity |
|------|------------|-------------|
| 入口文件 | `SKILL.md`（可嵌套在子 `skills/` 下） | `SKILL.md`（直接在 skill 文件夹根目录） |
| 元数据 | `name`, `description` | `name`(可选), `description`(必填) |
| 命令系统 | `commands/` 目录，每个 `.md` 文件是一个 slash command | ❌ 不支持，需合并到 SKILL.md |
| Hooks | `hooks/hooks.json` 定义 Session 钩子 | ❌ 不支持，需移除 |
| 设置文件 | `settings.json` | ❌ 不支持，需内联到 SKILL.md |
| 路径变量 | `${CLAUDE_PLUGIN_ROOT}` | 无等价变量，使用**相对于 SKILL.md 的路径** |
| 子目录约定 | `scripts/`, `skills/`(嵌套) | `scripts/`, `examples/`, `resources/` |
| frontmatter 特有字段 | `allowed-tools`, `argument-hint` | 无 |

## Proposed Changes

### 核心 SKILL.md 重写

完全重写 SKILL.md，整合以下内容：

1. **YAML frontmatter**：保留 `name` 和 `description`，去掉 Claude Code 特有字段
2. **合并 5 个 command 文件**的核心逻辑为 SKILL.md 中的使用场景章节
3. **合并 settings.json** 的默认配置为 SKILL.md 中的 "默认配置" 章节
4. **路径替换**：所有 `${CLAUDE_PLUGIN_ROOT}/scripts/` → 相对路径 `scripts/`
5. **合并嵌套 `skills/ppt-creation/SKILL.md`** 内容（API 速查、布局、配色等）

### 目录结构调整

- 删除 `commands/` 目录
- 删除 `hooks/` 目录
- 删除 `settings.json`
- 删除嵌套 `skills/` 目录（内容已合并到 SKILL.md）
- `templates/` → 移至 `resources/templates/`
- `references/` → 移至 `resources/`

### 最终目标目录结构

```
.agent/skills/pptxmaker/
├── SKILL.md
├── scripts/
│   ├── ppt_helpers.py
│   ├── create_ppt.py
│   ├── read_ppt.py
│   ├── check_deps.py
│   ├── requirements.txt
│   ├── _analyze_template.py
│   ├── _debug_shapes.py
│   └── _gen_report.py
├── resources/
│   ├── layouts.md
│   ├── color-palettes.md
│   ├── chart-guide.md
│   └── templates/
│       ├── README.md
│       ├── blank.pptx
│       ├── company-1.pptx
│       └── company-2026.pptx
└── examples/ (可选)
```
