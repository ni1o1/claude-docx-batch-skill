<p align="center">
  <img src="https://img.shields.io/badge/Claude%20Code-Skill-blue?style=for-the-badge" alt="Claude Code Skill">
  <img src="https://img.shields.io/badge/Word-Batch%20Edit-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white" alt="Word">
</p>

<h1 align="center">Claude DOCX Batch Skill</h1>

<p align="center">
  <strong>高效的 Word 文档批量编辑工具</strong>
  <br>
  <em>基于 python-docx 封装，一行代码批量修改格式</em>
</p>

<p align="center">
  <a href="#-features">Features</a> •
  <a href="#-quick-start">Quick Start</a> •
  <a href="#-usage">Usage</a> •
  <a href="#-vs-official">VS Official</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Claude%20Code-CLI-8A2BE2?logo=anthropic&logoColor=white" alt="Claude Code">
  <img src="https://img.shields.io/badge/Python-3.7+-3776AB?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-yellow" alt="License">
</p>

**English** | [中文](#中文)

---

## Overview

**Claude DOCX Batch Skill** provides efficient batch editing capabilities for Word documents. Unlike the official docx skill which manipulates raw XML, this skill uses python-docx for a simpler, more intuitive API.

### Why This Skill?

| Official DOCX Skill | This Skill |
|---------------------|------------|
| Unpack → Edit XML → Repack | Direct python-docx API |
| Complex XML manipulation | Simple `batch_update()` call |
| Good for tracked changes | Good for batch formatting |
| Steep learning curve | Easy to use |

---

## Features

| Feature | Description |
|---------|-------------|
| **Batch Formatting** | Font, alignment, indent, line spacing |
| **Paragraph Ops** | Insert, delete, replace text |
| **Table Ops** | Edit cells, rows, columns |
| **Image Ops** | Resize, delete, insert |
| **Global Replace** | Regex-supported text replacement |
| **Auto Numbering** | Clean up Word's auto numbering |

---

## Quick Start

### Installation

```bash
cd ~/.claude/skills
git clone https://github.com/YOUR_USERNAME/claude-docx-batch-skill.git
```

### Verify Installation

```bash
ls ~/.claude/skills/claude-docx-batch-skill/SKILL.md
```

### Dependencies

```bash
pip install python-docx
```

---

## Usage

### Basic Example

```python
from docx_editor import DocxEditor

editor = DocxEditor('input.docx')

# Get document outline
outline = editor.get_outline()
print(f"Total paragraphs: {outline['total']}")

# Batch update
editor.batch_update([
    {'op': 'update_style', 'index': 0,
     'font': {'name': '宋体'},
     'alignment': 'left',
     'spacing': {'line': 1.5, 'before': 0, 'after': 0}},
])

editor.save('output.docx')
```

### Format Entire Document

```python
editor = DocxEditor('thesis.docx')
outline = editor.get_outline()

ops = []
for i in range(outline['total']):
    ops.append({
        'op': 'update_style',
        'index': i,
        'alignment': 'left',
        'indent': {'first_line': 0, 'left': 0},
        'spacing': {'before': 0, 'after': 0, 'line': 1.15},
        'font': {'name': '宋体'}
    })

editor.batch_update(ops)
editor.save('thesis_formatted.docx')
```

---

## VS Official

| Aspect | Official docx skill | docx-batch |
|--------|---------------------|------------|
| **Method** | Unpack XML + DOM ops | python-docx wrapper |
| **Create docs** | docx-js (JavaScript) | Not supported |
| **Edit docs** | Manual XML editing | `batch_update()` |
| **Tracked changes** | ✅ Supported | ❌ Not supported |
| **Batch formatting** | Complex | ✅ Simple & efficient |
| **Use case** | Legal/contract review | Thesis/report formatting |

**Use this skill when:**
- You need to unify document formatting quickly
- Batch modify styles across many paragraphs
- Simple editing without tracked changes

**Use official skill when:**
- You need tracked changes (redlining)
- Working with legal/academic documents requiring revision history
- Creating new documents from scratch

---

## 中文

### 概述

**Claude DOCX Batch Skill** 提供高效的 Word 文档批量编辑能力。与官方 docx skill 操作原始 XML 不同，本 skill 使用 python-docx 提供更简洁直观的 API。

### 为什么用这个 Skill？

| 官方 DOCX Skill | 本 Skill |
|-----------------|----------|
| 解压 → 编辑 XML → 重新打包 | 直接调用 python-docx API |
| 复杂的 XML 操作 | 简单的 `batch_update()` 调用 |
| 适合 tracked changes | 适合批量格式化 |
| 学习曲线陡峭 | 容易上手 |

### 功能特性

| 功能 | 描述 |
|------|------|
| **批量格式化** | 字体、对齐、缩进、行距 |
| **段落操作** | 插入、删除、替换文本 |
| **表格操作** | 编辑单元格、行、列 |
| **图片操作** | 调整大小、删除、插入 |
| **全局替换** | 支持正则的文本替换 |
| **自动编号** | 清理 Word 自动编号 |

### 安装

```bash
cd ~/.claude/skills
git clone https://github.com/YOUR_USERNAME/claude-docx-batch-skill.git
```

### 依赖

```bash
pip install python-docx
```

### 使用示例

```python
from docx_editor import DocxEditor

editor = DocxEditor('论文.docx')
outline = editor.get_outline()

# 统一全文格式
ops = []
for i in range(outline['total']):
    ops.append({
        'op': 'update_style',
        'index': i,
        'alignment': 'left',
        'indent': {'first_line': 0, 'left': 0},
        'spacing': {'before': 0, 'after': 0, 'line': 1.15},
        'font': {'name': '宋体'}
    })

editor.batch_update(ops)
editor.save('论文_格式化.docx')
```

---

## API Reference

See [SKILL.md](skills/docx-batch/SKILL.md) for complete API documentation.

---

## Contributors

- **Claude** (Anthropic Claude Opus 4.5) - Skill Development

## License

MIT License - see [LICENSE](LICENSE) for details.

---

<p align="center">
  <sub>Built with collaboration between human and AI</sub>
</p>
