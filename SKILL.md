---
name: docx-batch
description: "高效的 Word 文档批量编辑工具，基于 python-docx 封装。适用于：(1) 批量格式化（字体、对齐、缩进、行距）(2) 段落/表格/图片的增删改 (3) 全局文本替换 (4) 清理自动编号。当需要快速统一文档格式、批量修改样式时使用此 skill，而非逐个操作 XML。"
license: MIT
---

# DocxEditor - Word 文档批量编辑工具

使用此 skill 前，确保已安装依赖：`pip install python-docx`

## 核心概念

### 三轨制索引

文档由三个独立列表组成，索引互不干扰：

| 轨道 | 索引参数 | 说明 |
|------|----------|------|
| 段落 | `index` | 包含标题和正文 |
| 表格 | `table_index` | 独立于段落 |
| 图片 | `image_index` | 独立于段落 |

### 工作流程

1. **查询** → 获取最新索引和内容
2. **构建操作列表** → 组装 `batch_update` 参数
3. **执行** → `batch_update` 自动倒序执行，无需手动计算索引漂移
4. **保存** → `save()` 写入文件

## 快速开始

```python
import sys
sys.path.insert(0, 'SKILL_DIR/scripts')  # 替换为实际路径
from docx_editor import DocxEditor

editor = DocxEditor('input.docx')
outline = editor.get_outline()
print(f"共 {outline['total']} 段落")

editor.batch_update([
    {'op': 'update_style', 'index': 0, 'font': {'name': '宋体'}, 'alignment': 'left'},
])
editor.save('output.docx')
```

## 查询接口

| 方法 | 用途 |
|------|------|
| `get_outline()` | 获取文档大纲（标题层级和索引） |
| `read_content(indices)` | 读取指定段落详情 |
| `get_tables_outline()` | 获取表格概览 |
| `read_table(table_index)` | 读取表格内容 |
| `get_images_outline()` | 获取图片概览 |

## 修改操作

所有修改通过 `batch_update(operations)` 执行，常用操作：

### 段落

```python
{'op': 'delete', 'index': 50}                    # 删除（自动保护含图片的段落）
{'op': 'delete', 'index': 50, 'force': True}     # 强制删除
{'op': 'insert', 'index': 10, 'position': 'after', 'text': '新内容'}
{'op': 'update_style', 'index': 20,
 'font': {'name': '宋体', 'size': 12, 'bold': False},
 'alignment': 'left',  # left/center/right/justify
 'indent': {'first_line': 0.74, 'left': 0},  # 厘米
 'spacing': {'before': 0, 'after': 0, 'line': 1.5}}  # line 是倍数
{'op': 'replace_text', 'index': 30, 'pattern': r'^（\d）\s*', 'replacement': '', 'regex': True}
{'op': 'replace_text_global', 'pattern': '旧文本', 'replacement': '新文本'}
{'op': 'clean_xml', 'index': 15, 'remove': ['numPr']}  # 清理自动编号
{'op': 'set_text', 'index': 25, 'text': '新的段落内容'}
```

### 表格

```python
{'op': 'update_table_cell', 'table_index': 0, 'row': 1, 'col': 2, 'text': '新内容'}
{'op': 'update_table_row', 'table_index': 0, 'row': 1, 'texts': ['列1', '列2', '列3']}
{'op': 'update_table_col', 'table_index': 0, 'col': 1, 'texts': ['行1', '行2', '行3']}
```

### 图片

```python
{'op': 'delete_image', 'image_index': 0}
{'op': 'resize_image', 'image_index': 0, 'width': 10.0}  # 宽度cm，高度自动
{'op': 'insert_image', 'index': 10, 'path': '/path/to/img.png', 'width': 6.0}
```

### 系统

```python
{'op': 'update_fields_on_open'}  # 让Word打开时刷新目录/页码
```

## 重要：删除操作的安全保护

Word 段落可能包含隐藏内容（图片、OLE对象），`text` 属性为空但实际有内容。

`read_content` 返回的关键字段：

| 字段 | 含义 |
|------|------|
| `is_empty` | 文本为空（可能包含图片） |
| `is_truly_empty` | 真正为空，可安全删除 |
| `has_embedded` | 是否包含图片/OLE等 |

安全清理空段落：
```python
content = editor.read_content(range(100))
ops = [{'op': 'delete', 'index': p['index']} for p in content if p['is_truly_empty']]
editor.batch_update(ops)
```

## 限制

- 不支持 tracked changes（需要请用官方 docx skill）
- 不支持读取图片内容，只能操作尺寸或删除
- 目录/页码需用 `update_fields_on_open`，让 Word 自己计算
