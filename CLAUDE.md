# 公文写作 Skill 开发

## 项目简介

本项目开发一个 Claude Code skill（`gongwen-writing`），用于将用户输入的普通文本、Word文档或PDF文档改写/撰写为符合中国党政机关公文规范的正式文本，同时不违反任何法律法规。

## 设计文档

- 设计规格：`docs/superpowers/specs/2026-03-19-gongwen-writing-design.md`

## 技术方案

- **纯 Prompt Skill**：单个 SKILL.md 文件，内嵌公文规范和红线规则
- **输入**：纯文本 / Word / PDF
- **输出**：终端显示 或 导出 Word（用户选择）
- **规范基础**：GB/T 9704-2012 + 《党政机关公文处理工作条例》
- **文种覆盖**：15种法定公文 + 事务性文书
- **依赖**：`office-mcp`、`word-document-processor` skill

## 红线要求

- 禁止伪造公文（虚构发文机关、文号、印章）
- 遵守新华社禁用词/慎用词规范
- 敏感话题（领导人、民族宗教、领土主权）审慎处理
- 每次输出附带AI生成免责声明
