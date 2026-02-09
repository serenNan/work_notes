---
tags:
  - claude-code
  - agent-skills
  - AI工具
  - 教程
source: "https://www.bilibili.com/video/BV1G3FNznEiS/"
date: 2026-02-09
type: tech
author: 秋芝2046
---

# 手把手彻底学会 Agent Skills

> [!note] 视频信息
> UP主：秋芝2046 | 时长：19分19秒 | 播放：17.5w | 收藏：1.3w
> 一次搞懂 Agent Skills 原理，手把手带你做出自己的 Skill

## 概述

Agent Skills 是将专业知识、工作流规范固化为**可复用资产**的核心工具。本质上是一个模块化的 Markdown 文件（SKILL.md），能教会 AI 工具执行特定任务，支持自动触发、团队共享与工程化管理，彻底告别重复的提示词输入。

Agent Skills 是由 Anthropic 牵头维护的**开放标准**，目前已得到 Anthropic / OpenAI / Google / Microsoft / Cursor 等多家公司支持，迅速成为各大主流 AI 工具的标配。

## 核心原理：渐进式披露（Progressive Disclosure）

> [!tip] Skills 与传统 Prompt 的最大区别
> **按需加载 + 渐进式披露**——只在需要时才把详细的 SOP 塞进上下文，极大节省 token。

Agent Skills 分**三层**加载：

```
层级 1：技能发现    → AI 读取所有技能的 name + description（始终在上下文中）
        ↓ 判断是否相关
层级 2：加载核心指令  → 自动读取 SKILL.md 正文，获取详细指导
        ↓ 需要更多细节时
层级 3：加载资源文件  → 按需读取额外文件（脚本、示例、参考文档）
```

| 层级 | 加载内容 | 何时加载 | token 开销 |
|------|----------|----------|------------|
| 1 | name + description | 始终 | 极低 |
| 2 | SKILL.md 正文 | 技能被调用时 | 中等 |
| 3 | 支持文件/脚本 | 明确需要时 | 按需 |

## SKILL.md 文件结构

每个 Skill 的核心就是一个 `SKILL.md` 文件，由两部分组成：

```yaml
---
# ===== YAML 前置元数据（frontmatter）=====
name: your-skill-name          # 技能名称，变成 /slash-command
description: 技能做什么，何时使用  # Claude 根据这个判断是否自动加载
---

# ===== Markdown 正文（指令内容）=====

## 指令
具体的执行步骤和规则...

## 示例
使用示例...

## 注意事项
容易踩坑的地方...
```

## 前置元数据（Frontmatter）详解

| 字段 | 必需 | 说明 |
|------|------|------|
| `name` | 否 | 技能名称，仅小写字母、数字和连字符（最多64字符）。省略则用目录名 |
| `description` | **推荐** | 技能描述。Claude 据此决定何时自动使用 |
| `argument-hint` | 否 | 自动补全提示，如 `[issue-number]` |
| `disable-model-invocation` | 否 | 设为 `true` 禁止 Claude 自动调用，只能手动 `/name` 触发 |
| `user-invocable` | 否 | 设为 `false` 从 `/` 菜单隐藏，仅 Claude 自动使用 |
| `allowed-tools` | 否 | 限制技能可使用的工具，如 `Read, Grep, Glob` |
| `model` | 否 | 指定技能使用的模型 |
| `context` | 否 | 设为 `fork` 在隔离的子代理中运行 |
| `agent` | 否 | 搭配 `context: fork` 指定子代理类型（`Explore`、`Plan` 等） |
| `hooks` | 否 | 技能生命周期钩子 |

### 调用控制矩阵

| 前置元数据设置 | 用户可调用 | Claude 可自动调用 | 适用场景 |
|----------------|-----------|-------------------|----------|
| （默认） | ✅ | ✅ | 通用技能 |
| `disable-model-invocation: true` | ✅ | ❌ | 部署、提交等有副作用的操作 |
| `user-invocable: false` | ❌ | ✅ | 背景知识，不适合作为命令 |

## 技能存放位置与作用范围

| 位置 | 路径 | 适用范围 |
|------|------|----------|
| 个人全局 | `~/.claude/skills/<skill-name>/SKILL.md` | 你的所有项目 |
| 项目专用 | `.claude/skills/<skill-name>/SKILL.md` | 仅当前项目 |
| 插件 | `<plugin>/skills/<skill-name>/SKILL.md` | 启用插件的地方 |
| 企业级 | 通过托管设置部署 | 组织内所有用户 |

> [!warning] 注意
> 项目技能会**覆盖**同名的个人技能。如果不想提交到 git，把 `.claude` 加到 `.gitignore`。

### 目录结构

```
my-skill/
├── SKILL.md           # 主要说明（必需）
├── reference.md       # 详细参考文档（按需加载）
├── examples/
│   └── sample.md      # 示例输出
└── scripts/
    └── helper.py      # Claude 可执行的脚本
```

> [!tip] 最佳实践
> 保持 `SKILL.md` 在 **500 行以下**，详细参考材料放到单独文件并在 SKILL.md 中引用链接。

## 实战：创建你的第一个 Skill

### 步骤 1：创建目录

```bash
mkdir -p ~/.claude/skills/explain-code
```

### 步骤 2：编写 SKILL.md

```yaml
---
name: explain-code
description: Explains code with visual diagrams and analogies. Use when explaining how code works.
---

When explaining code, always include:

1. **Start with an analogy**: Compare the code to something from everyday life
2. **Draw a diagram**: Use ASCII art to show the flow
3. **Walk through the code**: Explain step-by-step
4. **Highlight a gotcha**: Common mistake or misconception?

Keep explanations conversational.
```

### 步骤 3：测试

```
# 方式一：自然语言触发（Claude 自动匹配）
How does this code work?

# 方式二：斜杠命令直接调用
/explain-code src/auth/login.ts
```

### 更多示例

**部署技能（仅手动触发）：**

```yaml
---
name: deploy
description: Deploy the application to production
disable-model-invocation: true   # 防止 Claude 自作主张部署
---

Deploy $ARGUMENTS to production:
1. Run the test suite
2. Build the application
3. Push to the deployment target
4. Verify the deployment succeeded
```

**只读浏览技能：**

```yaml
---
name: safe-reader
description: Read files without making changes
allowed-tools: Read, Grep, Glob   # 限制只能读不能写
---
```

## 高级功能

### 1. 参数传递 `$ARGUMENTS`

技能名称后面的所有内容都会替换 `$ARGUMENTS`：

```yaml
---
name: fix-issue
description: Fix a GitHub issue
---

Fix GitHub issue $ARGUMENTS following our coding standards.
```

运行 `/fix-issue 123` → Claude 收到 "Fix GitHub issue **123** following our coding standards."

### 2. 动态上下文注入 `` !`command` ``

在技能内容发送给 Claude **之前**执行 shell 命令，输出替换占位符：

```yaml
---
name: pr-summary
description: Summarize changes in a pull request
context: fork
agent: Explore
---

## Pull request context
- PR diff: !`gh pr diff`
- PR comments: !`gh pr view --comments`

## Your task
Summarize this pull request...
```

> [!note] 这是预处理
> 命令在 Claude 看到内容之前就已执行完毕，Claude 只看到最终结果。

### 3. 子代理运行 `context: fork`

添加 `context: fork` 让技能在隔离环境中运行，不共享对话历史：

```yaml
---
name: deep-research
description: Research a topic thoroughly
context: fork
agent: Explore        # 使用 Explore 代理（只读、擅长搜索）
---

Research $ARGUMENTS thoroughly:
1. Find relevant files using Glob and Grep
2. Read and analyze the code
3. Summarize findings with specific file references
```

### 4. 生成视觉输出

技能可以捆绑脚本生成交互式 HTML 报告，在浏览器中打开：

```yaml
---
name: codebase-visualizer
description: Generate an interactive tree visualization of your codebase
allowed-tools: Bash(python:*)
---

Run the visualization script:
python ~/.claude/skills/codebase-visualizer/scripts/visualize.py .
```

## Skills vs 其他概念对比

| 特性 | Skills | CLAUDE.md | Hooks | MCP Tools | Sub-agents |
|------|--------|-----------|-------|-----------|------------|
| 触发方式 | 用户调用 / Claude 自动 | 始终加载 | 事件驱动 | Claude 按需 | Claude 委派 |
| token 效率 | 高（渐进式） | 低（全量加载） | 不占 token | 中等 | 独立上下文 |
| 可复用性 | ⭐⭐⭐ | ⭐ | ⭐⭐ | ⭐⭐⭐ | ⭐⭐ |
| 适用场景 | 可复用工作流 | 全局规则 | 前/后置处理 | 外部系统集成 | 复杂独立任务 |

## 故障排除

| 问题 | 解决方案 |
|------|----------|
| 技能未触发 | 检查 description 是否包含用户会自然说的关键字；用 `/skill-name` 直接调用测试 |
| 技能触发过于频繁 | 让 description 更具体；添加 `disable-model-invocation: true` |
| Claude 看不到所有技能 | 运行 `/context` 检查；设置 `SLASH_COMMAND_TOOL_CHAR_BUDGET` 增加限制 |

## 总结与要点回顾

- [ ] Agent Skills 本质是**模块化 Markdown 文件**，核心是 SKILL.md
- [ ] 原理是**渐进式披露**：metadata → 指令 → 资源，按需加载省 token
- [ ] 存放位置决定作用范围：个人全局 `~/.claude/skills/` vs 项目专用 `.claude/skills/`
- [ ] `description` 字段最重要——决定 Claude 何时自动使用技能
- [ ] `disable-model-invocation: true` 防止 Claude 自动触发有副作用的操作
- [ ] `context: fork` 让技能在隔离子代理中运行
- [ ] `` !`command` `` 语法可在加载前注入动态数据
- [ ] 保持 SKILL.md 精简（<500行），详细内容放支持文件

> [!warning] 常见误区
> - Skill 文件名必须是 `SKILL.md`（全大写 + .md 小写），不是随意命名
> - `context: fork` 的技能必须有明确的任务指令，不能只放指南
> - `user-invocable: false` 只控制菜单可见性，不阻止程序化调用

## 扩展阅读

- [Agent Skills 官方文档（中文）](https://code.claude.com/docs/zh-CN/skills)
- [Agent Skills 开放标准](https://agentskills.io)
- [awesome-agent-skills 终极指南](https://github.com/libukai/awesome-agent-skills)
- [AgentSkills.best 市场](https://agentskills.best/)
- [菜鸟教程 - Agent Skills](https://www.runoob.com/claude-code/claude-agent-skills.html)
