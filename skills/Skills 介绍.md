---
title: Claude Code Skills 介绍
date: 2026-02-09
tags: [claude-code, skills, 工具]
---

# Claude Code Skills 介绍

## 什么是 Skills

Skills 是 Claude Code 中的**可复用能力模块**, 类似于"斜杠命令" (`/command`). 每个 Skill 封装了一组特定的操作流程, 用户可以通过简短的指令触发复杂的自动化工作流.

> [!tip] 核心价值
> Skills 让你把**重复性的多步骤任务**打包成一个命令, 一键执行.

## Skills 的工作原理

```
用户输入 /skill-name  →  Claude 识别并加载 Skill  →  按预设流程执行  →  返回结果
```

1. **定义**: 每个 Skill 在配置中声明名称, 描述和执行逻辑
2. **触发**: 用户通过 `/skill-name` 或自然语言描述触发
3. **执行**: Claude 按照 Skill 内置的 prompt 和工具链自动完成任务

## 当前可用 Skills

| Skill 名称 | 说明 | 使用场景 |
|---|---|---|
| `doc-imple` | 根据文档实现代码 | 有设计文档, 需要自动生成对应代码 |
| `explain` | 运行通知脚本 | 任务完成后发送通知 |
| `markdown` | 文档格式化 | 整理和美化 Markdown 文档 |
| `myreview` | Git 项目概览 | 快速了解 Git 仓库的整体状态 |
| `project-read` | 历史对话摘要读取 | 读取之前保存的对话上下文 |
| `project-summary` | 写对话摘要文档 | 将当前对话整理为 `历史对话摘要.md` |
| `workflow` | 完整开发工作流提示词 | 启动标准化的开发流程 |
| `claude-code-setup` | 自动化推荐分析 | 分析代码库并推荐 Claude Code 自动化配置 |

## 如何创建自定义 Skill

Skills 的定义文件通常存放在项目的 `.claude/skills/` 目录下, 格式为 Markdown 文件:

```
.claude/
└── skills/
    ├── my-skill.md
    └── another-skill.md
```

### Skill 文件结构示例

```markdown
---
name: my-custom-skill
description: 这个 Skill 做什么
---

# 执行步骤

1. 第一步: 读取某些文件
2. 第二步: 分析内容
3. 第三步: 生成输出
```

> [!note] 关键要点
> - Skill 文件本质上是一段**结构化的 prompt**
> - 它告诉 Claude 在被触发时应该执行哪些步骤
> - 可以引用工具 (Read, Write, Bash 等) 来完成具体操作

## 使用方式

### 方式一: 斜杠命令

在 Claude Code 中直接输入:

```
/skill-name
/skill-name 参数
```

### 方式二: 自然语言

直接描述你的需求, Claude 会自动匹配合适的 Skill:

```
帮我格式化这个文档        → 触发 markdown skill
给我看看这个项目的概览     → 触发 myreview skill
```

## Skills vs 其他概念对比

| 特性 | Skills | Hooks | MCP Tools |
|---|---|---|---|
| 触发方式 | 用户主动调用 | 自动触发 (事件驱动) | Claude 按需调用 |
| 定义位置 | `.claude/skills/` | `.claude/hooks/` | MCP 服务器配置 |
| 复杂度 | 中等 | 低 | 高 |
| 适用场景 | 可复用工作流 | 前/后置处理 | 外部系统集成 |

> [!warning] 注意事项
> - Skills 的执行依赖当前项目上下文, 确保在正确的项目目录下使用
> - 自定义 Skill 的 prompt 质量直接影响执行效果
> - 建议为每个 Skill 写清晰的 description, 方便自动匹配
