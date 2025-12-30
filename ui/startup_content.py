# -*- coding: utf-8 -*-

from __future__ import annotations


def insert_example_if_empty_for_app(app) -> None:
    """仅在编辑器为空时插入示例文本，避免打扰用户已有内容。"""
    try:
        current = app.input_text.get('1.0', 'end-1c')
        if (current or '').strip():
            return
    except Exception:
        pass

    example = """# 欢迎使用 Markdown 转换器 

## 核心功能

这是一个**功能完善**的 Markdown 转 Word 工具：

### 文档转换
- ✅ 标题、段落、列表（有序/无序）
- ✅ **粗体**、*斜体*、~~删除线~~
- ✅ 上标<sup>2</sup>和下标<sub>2</sub>
- ✅ 表格（自动三线表样式）
- ✅ 数学公式（LaTeX 语法）
- ✅ 代码块高亮
- ✅ 图片自动缩放
- ✅ 可点击超链接

### 任务列表
- [ ] 待完成任务
- [x] 已完成任务

### 编辑功能
- ✅ 保存源文件（Ctrl+S）
- ✅ 导出Word（Ctrl+Shift+S）
- ✅ 撤销/重做（Ctrl+Z / Ctrl+Y）
- ✅ 查找/替换（Ctrl+F / Ctrl+H）
- ✅ 未保存提示

### 界面特性
- ✅ 实时预览
- ✅ 亮/暗主题切换
- ✅ 窗口位置记忆
- ✅ 最近文件列表


### Phase 2 新增功能 (New!)
- 🎨 **主题编辑器**: 自定义界面颜色，打造个性化外观
- 📄 **Word模板库**: 导入 .docx 模板，一键生成红头文件/企业文档
- 📝 **页眉页脚**: 自定义页眉文字、日期及页码
- 📎 **脚注支持**: 自动解析 `[^1]` 并生成尾注

## 数学公式示例

行内公式：质能方程 $E = mc^2$

块级公式：

$$
\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}
$$

## 代码示例

```python
def hello():
    print("Hello, World!")
```

## 脚注示例

Markdown 转换器支持脚注[^1]功能了。

[^1]: 这是一个脚注的示例内容。

## 快捷键

| 功能 | 快捷键 |
|------|--------|
| 保存源文件 | Ctrl+S |
| 导出Word | Ctrl+Shift+S |
| 打开文件 | Ctrl+O |
| 撤销 | Ctrl+Z |
| 重做 | Ctrl+Y |
| 查找 | Ctrl+F |
| 替换 | Ctrl+H |
| 帮助 | F1 |
"""

    try:
        app.input_text.insert('1.0', example)
    except Exception:
        return

    try:
        app.on_text_change(None)
    except Exception:
        pass
