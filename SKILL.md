---
name: docx-batch
description: "高效的 Word 文档批量编辑工具，基于 python-docx 封装。适用于：(1) 批量格式化（字体、对齐、缩进、行距）(2) 段落/表格/图片的增删改 (3) 全局文本替换 (4) 清理自动编号。当需要快速统一文档格式、批量修改样式时使用此 skill，而非逐个操作 XML。"
---

# DocxEditor - Word 文档批量编辑工具

## 核心思维模型

### 三轨制索引

文档由三个独立列表组成，索引互不干扰：

| 轨道 | 索引参数 | 说明 |
|------|----------|------|
| 段落 | `index` | 包含标题和正文 |
| 表格 | `table_index` | 独立于段落 |
| 图片 | `image_index` | 独立于段落 |

### 无状态执行

- 每次修改前先调用查询接口获取最新索引
- `batch_update` 内部自动倒序执行，无需手动计算索引漂移

## 快速开始

```python
from docx_editor import DocxEditor

editor = DocxEditor('input.docx')

# 查看文档结构
outline = editor.get_outline()
print(f"共 {outline['total']} 段落")

# 批量修改
editor.batch_update([
    {'op': 'update_style', 'index': 0, 'font': {'name': '宋体'}, 'alignment': 'left'},
    {'op': 'update_style', 'index': 1, 'spacing': {'line': 1.5, 'before': 0, 'after': 0}},
])

editor.save('output.docx')
```

## 查询接口

| 方法 | 用途 |
|------|------|
| `get_outline()` | 获取文档大纲（标题层级和索引） |
| `read_content(indices)` | 读取指定段落详情（文本、样式、XML属性） |
| `get_tables_outline()` | 获取表格概览 |
| `read_table(table_index)` | 读取表格内容 |
| `get_images_outline()` | 获取图片概览 |

## 修改操作 (`batch_update`)

### 段落操作

```python
# 删除段落（自动保护含图片/OLE对象的段落）
{'op': 'delete', 'index': 50}

# 强制删除（跳过保护检查，慎用）
{'op': 'delete', 'index': 50, 'force': True}

# 插入段落
{'op': 'insert', 'index': 10, 'position': 'after', 'text': '新内容', 'style': 'Normal'}

# 修改样式（最常用）
{'op': 'update_style', 'index': 20,
 'style': 'Normal',
 'font': {'name': '宋体', 'size': 12, 'bold': False},
 'alignment': 'left',  # left/center/right/justify
 'indent': {'first_line': 0.74, 'left': 0},  # 单位：厘米
 'spacing': {'before': 0, 'after': 0, 'line': 1.5}  # line 是倍数
}

# 替换文本（支持正则）
{'op': 'replace_text', 'index': 30, 'pattern': r'^（\d）\s*', 'replacement': '', 'regex': True}

# 全局替换
{'op': 'replace_text_global', 'pattern': '旧文本', 'replacement': '新文本', 'regex': False}

# 清理自动编号（重要！）
{'op': 'clean_xml', 'index': 15, 'remove': ['numPr']}

# 设置段落文本
{'op': 'set_text', 'index': 25, 'text': '新的段落内容'}
```

### 表格操作

```python
# 修改单元格
{'op': 'update_table_cell', 'table_index': 0, 'row': 1, 'col': 2, 'text': '新内容'}

# 批量修改整行
{'op': 'update_table_row', 'table_index': 0, 'row': 1, 'texts': ['列1', '列2', '列3']}

# 批量修改整列
{'op': 'update_table_col', 'table_index': 0, 'col': 1, 'texts': ['行1', '行2', '行3']}
```

### 图片操作

```python
# 删除图片
{'op': 'delete_image', 'image_index': 0}

# 调整大小（宽度cm，高度自动保持比例）
{'op': 'resize_image', 'image_index': 0, 'width': 10.0}

# 插入图片
{'op': 'insert_image', 'index': 10, 'path': '/path/to/img.png', 'width': 6.0}
```

### 系统操作

```python
# 刷新目录/页码/交叉引用（让Word打开时自动更新）
{'op': 'update_fields_on_open'}
```

## 重要注意事项

### 段落可能包含隐藏内容

Word 段落中除了可见文本，还可能包含：
- 图片（inline shape）
- 换页符（page break）
- 分节符（section break）
- 其他嵌入对象

这些内容存在于 `runs` 中，但 **不会体现在 `text` 属性里**。

### 删除操作的安全保护

`delete` 操作会**自动保护**包含嵌入对象（图片、OLE对象、图表等）的段落：

- 这类段落的 `text` 属性为空，但实际包含重要内容
- 删除时会抛出 `ValueError`，防止误删
- 如确需删除，使用 `force: True` 强制执行

`read_content` 返回的字段说明：

| 字段 | 含义 |
|------|------|
| `is_empty` | 文本为空（可能包含图片） |
| `is_truly_empty` | 真正为空（无文字且无嵌入对象），可安全删除 |
| `has_embedded` | 是否包含图片/OLE等嵌入对象 |

```python
# 示例：安全清理空段落
content = editor.read_content(range(100))
ops = [{'op': 'delete', 'index': p['index']} for p in content if p['is_truly_empty']]
editor.batch_update(ops)  # 只删除真正的空段落，自动跳过含图片的
```

## 常见场景示例

### 统一全文格式

```python
editor = DocxEditor('论文.docx')
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
editor.save('论文_格式化.docx')
```

### 清理自动编号

如果 `read_content` 显示 `has_numPr: true`，必须用 `clean_xml` 移除：

```python
content = editor.read_content([10, 11, 12])
ops = []
for p in content:
    if p['xml']['has_numPr']:
        ops.append({'op': 'clean_xml', 'index': p['index'], 'remove': ['numPr']})
editor.batch_update(ops)
```

## 限制

| 禁止事项 | 原因 |
|----------|------|
| 不要手动更新页码/目录 | 用 `update_fields_on_open`，让 Word 自己计算 |
| 不支持 tracked changes | 需要复杂编辑请用官方 docx skill |
| 不支持读取图片内容 | 只能操作图片尺寸或删除 |

## 依赖

```bash
pip install python-docx
```

## 脚本位置

核心脚本：`scripts/docx_editor.py`
