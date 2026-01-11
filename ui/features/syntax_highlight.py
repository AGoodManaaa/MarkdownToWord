# -*- coding: utf-8 -*-
"""Markdown 智能语法高亮模块"""

import re
import tkinter as tk
import tkinter.font as tkfont
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


@dataclass
class HighlightTheme:
    """语法高亮主题"""
    heading: str = "#3b82f6"      # 标题 - 蓝色
    heading_marker: str = "#9ca3af"  # # 符号 - 灰色
    bold: str = "#1f2937"         # 粗体 - 深色
    italic: str = "#6b7280"       # 斜体 - 灰色
    code_inline: str = "#dc2626"  # 行内代码 - 红色
    code_block: str = "#059669"   # 代码块 - 绿色
    code_bg: str = "#f3f4f6"      # 代码背景 - 浅灰
    link: str = "#2563eb"         # 链接 - 蓝色
    link_url: str = "#9ca3af"     # URL - 灰色
    image: str = "#7c3aed"        # 图片 - 紫色
    list_marker: str = "#f59e0b"  # 列表符号 - 橙色
    blockquote: str = "#6b7280"   # 引用 - 灰色
    blockquote_bg: str = "#f9fafb"  # 引用背景
    hr: str = "#d1d5db"           # 分隔线 - 浅灰
    strikethrough: str = "#9ca3af"  # 删除线 - 灰色
    table: str = "#0891b2"        # 表格 - 青色
    math: str = "#7c3aed"         # 数学公式 - 紫色
    comment: str = "#9ca3af"      # 注释 - 灰色


# 默认主题
DEFAULT_THEME = HighlightTheme()

# Markdown 语法正则表达式
PATTERNS = {
    # 标题 (# ## ### 等) - 允许前置空格，并兼容：
    # 1) ### 标题（有空格）
    # 2) ###标题（无空格）
    # 3) 全角/不间断空格
    'heading': r'^[\s\u00a0\u3000]*(#{1,6})[\s\u00a0\u3000]*(.+?)\s*$',
    # 粗体 **text** 或 __text__
    'bold': r'(\*\*|__)(.+?)\1',
    # 斜体 *text* 或 _text_
    'italic': r'(?<!\*)(\*|_)(?!\*)(.+?)\1(?!\*)',
    # 粗斜体 ***text***
    'bold_italic': r'(\*\*\*|___)(.+?)\1',
    # 行内代码 `code`
    'code_inline': r'`([^`\n]+)`',
    # 代码块开始/结束 ```（允许前后空白、语言标识）
    'code_block_start': r'^\s*```\s*\w*\s*$',
    'code_block_end': r'^\s*```\s*$',
    # 链接 [text](url)
    'link': r'\[([^\]]+)\]\(([^)]+)\)',
    # 图片 ![alt](url)
    'image': r'!\[([^\]]*)\]\(([^)]+)\)',
    # 无序列表 - * +
    'unordered_list': r'^(\s*)([-*+])\s+',
    # 有序列表 1. 2. 等
    'ordered_list': r'^(\s*)(\d+\.)\s+',
    # 引用 >
    'blockquote': r'^(>\s*)+',
    # 分隔线 --- *** ___
    'hr': r'^([-*_]){3,}\s*$',
    # 删除线 ~~text~~
    'strikethrough': r'~~(.+?)~~',
    # 表格分隔符 |---|
    'table_separator': r'^\|?[\s:-]+\|[\s|:-]+\|?$',
    # 表格行 | cell |
    'table_row': r'^\|(.+)\|$',
    # 数学公式 $formula$ 或 $$formula$$
    'math_inline': r'\$([^$\n]+)\$',
    'math_block': r'\$\$([^$]+)\$\$',
    # 任务列表 - [ ] 或 - [x]
    'task_list': r'^(\s*)([-*+])\s+\[([ xX])\]\s+',
}


class SyntaxHighlighter:
    """Markdown 语法高亮器 - 支持虚拟化渲染"""
    
    def __init__(self, text_widget, theme: HighlightTheme = None, buffer_lines: int = 50):
        """
        初始化语法高亮器
        
        Args:
            text_widget: tkinter Text 或 CTkTextbox 组件
            theme: 高亮主题
            buffer_lines: 可视区域上下缓冲行数，用于虚拟化渲染
        """
        self.text_widget = text_widget
        self.theme = theme or DEFAULT_THEME
        self._enabled = True
        self._in_code_block = False
        self._code_block_start_line = 0
        self._debounce_id = None
        self._debounce_delay = 50  # ms
        
        # 虚拟化渲染参数
        self.buffer_lines = buffer_lines
        self._rendered_range = (0, 0)  # 已渲染的行范围
        self._last_visible_range = (0, 0)  # 上次可见范围
        self._highlight_cache = {}  # 行高亮缓存: {line_num: content_hash}
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 配置标签样式
        self._configure_tags()
        
        # 绑定事件
        self._bind_events()
        
        # 绑定滚动事件以支持虚拟化
        self._bind_scroll_events()

        # 初始全量高亮，避免部分行未刷新
        try:
            self.highlight_all()
        except Exception:
            pass
    
    def _configure_tags(self):
        """配置文本标签样式"""
        t = self._text
        theme = self.theme
        # 读取当前编辑器基础字体，确保粗体/斜体字号与正文一致
        base_font = tkfont.Font(font=t.cget('font'))
        base_family = base_font.actual('family')
        base_size = base_font.actual('size') or 16
        code_size = max(12, int(base_size) - 1)
        
        # 标题样式
        for i in range(1, 7):
            size = 24 - (i - 1) * 2  # h1=24, h2=22, ..., h6=14
            t.tag_configure(f'heading{i}', foreground=theme.heading,
                           font=('Microsoft YaHei', size, 'bold'))
            t.tag_configure(f'heading{i}_marker', foreground=theme.heading_marker)
        
        # 粗体
        t.tag_configure('bold', font=(base_family, base_size, 'bold'))
        t.tag_configure('bold_marker', foreground=theme.heading_marker)
        
        # 斜体
        t.tag_configure('italic', font=(base_family, base_size, 'italic'),
                       foreground=theme.italic)
        t.tag_configure('italic_marker', foreground=theme.heading_marker)
        
        # 粗斜体
        t.tag_configure('bold_italic', font=(base_family, base_size, 'bold italic'))
        
        # 行内代码
        t.tag_configure('code_inline', foreground=theme.code_inline,
                       background=theme.code_bg, font=('Consolas', code_size))
        
        # 代码块
        t.tag_configure('code_block', foreground=theme.code_block,
                       background=theme.code_bg, font=('Consolas', code_size))
        t.tag_configure('code_block_marker', foreground=theme.heading_marker,
                       background=theme.code_bg)
        
        # 链接
        t.tag_configure('link_text', foreground=theme.link, underline=True)
        t.tag_configure('link_url', foreground=theme.link_url)
        t.tag_configure('link_bracket', foreground=theme.heading_marker)
        
        # 图片
        t.tag_configure('image', foreground=theme.image)
        t.tag_configure('image_marker', foreground=theme.heading_marker)
        
        # 列表
        t.tag_configure('list_marker', foreground=theme.list_marker,
                       font=(base_family, base_size, 'bold'))
        
        # 引用
        t.tag_configure('blockquote', foreground=theme.blockquote,
                       background=theme.blockquote_bg,
                       lmargin1=20, lmargin2=20)
        t.tag_configure('blockquote_marker', foreground=theme.list_marker)
        
        # 分隔线
        t.tag_configure('hr', foreground=theme.hr)
        
        # 删除线
        t.tag_configure('strikethrough', overstrike=True,
                       foreground=theme.strikethrough)
        
        # 表格
        t.tag_configure('table', foreground=theme.table)
        t.tag_configure('table_separator', foreground=theme.hr)
        
        # 数学公式
        t.tag_configure('math', foreground=theme.math,
                       font=('Cambria Math', base_size))
        
        # 任务列表
        t.tag_configure('task_done', overstrike=True,
                       foreground=theme.strikethrough)
        t.tag_configure('task_checkbox', foreground=theme.list_marker)
    
    def _bind_events(self):
        """绑定文本变化事件"""
        self._text.bind('<KeyRelease>', self._on_key_release)
        self._text.bind('<<Modified>>', self._on_modified)
    
    def _bind_scroll_events(self):
        """绑定滚动事件以刷新可见区域高亮"""
        try:
            self._text.bind('<MouseWheel>', self._on_scroll, add='+')
            self._text.bind('<Button-4>', self._on_scroll, add='+')  # Linux up
            self._text.bind('<Button-5>', self._on_scroll, add='+')  # Linux down
        except Exception:
            pass
    
    def _on_key_release(self, event=None):
        """按键释放事件"""
        if not self._enabled:
            return
        
        # 防抖处理
        if self._debounce_id:
            self._text.after_cancel(self._debounce_id)
        
        self._debounce_id = self._text.after(
            self._debounce_delay,
            self._highlight_visible
        )
    
    def _on_modified(self, event=None):
        """文本修改事件"""
        if self._text.edit_modified():
            self._text.edit_modified(False)
            self._on_key_release()
    
    def _highlight_visible(self):
        """只高亮可见区域"""
        try:
            # 获取可见区域
            first_visible = self._text.index("@0,0")
            last_visible = self._text.index(f"@0,{self._text.winfo_height()}")
            
            # 扩展范围以包含完整的代码块
            first_line = int(first_visible.split('.')[0])
            last_line = int(last_visible.split('.')[0])
            
            # 向前扩展 10 行，向后扩展 10 行
            first_line = max(1, first_line - 10)
            last_line = last_line + 10
            
            self._highlight_range(f"{first_line}.0", f"{last_line}.end")
        except Exception:
            pass
    
    def highlight_all(self):
        """高亮全部内容"""
        if not self._enabled:
            return
        self._highlight_range("1.0", "end")
    
    def _highlight_range(self, start: str, end: str):
        """高亮指定范围"""
        # 清除现有标签（保持 sel）
        for tag in self._text.tag_names():
            if tag != 'sel':
                self._text.tag_remove(tag, start, end)
        # 额外确保 code_block 不延续到新渲染范围外
        self._in_code_block = False
        
        # 获取文本内容
        content = self._text.get(start, end)
        lines = content.split('\n')
        
        # 计算起始行号
        start_line = int(start.split('.')[0])
        
        # 逐行处理
        self._in_code_block = False

        for i, line in enumerate(lines):
            line_num = start_line + i
            line_start = f"{line_num}.0"
            line_end = f"{line_num}.end"
            
            stripped = line.strip()
            is_fence = stripped.startswith("```")
            # 检查代码块状态（允许行首/行内有空白，strip 后以 ``` 开头视为开始/结束）
            if re.match(PATTERNS['code_block_start'], line) or is_fence:
                # 如果已在代码块内且再次遇到围栏，先收尾再继续，防止漏关
                if self._in_code_block:
                    self._text.tag_add('code_block_marker', line_start, line_end)
                    self._in_code_block = False
                    continue
                self._in_code_block = True
                self._code_block_start_line = line_num
                self._text.tag_add('code_block_marker', line_start, line_end)
                continue
            
            if self._in_code_block and (re.match(PATTERNS['code_block_end'], line) or is_fence):
                self._in_code_block = False
                self._text.tag_add('code_block_marker', line_start, line_end)
                continue
            
            if self._in_code_block:
                self._text.tag_add('code_block', line_start, line_end)
                continue
            
            # 高亮当前行
            self._highlight_line(line, line_num)
    
    def _highlight_line(self, line: str, line_num: int):
        """高亮单行"""
        line_start = f"{line_num}.0"
        
        # 标题：改用“前导空白 + 连续# + 文本”判定，兼容无空格/全角空格/不间断空格
        stripped = line.lstrip(" \t\u00a0\u3000")
        if stripped.startswith('#'):
            level = 0
            for ch in stripped:
                if ch == '#':
                    level += 1
                else:
                    break
            level = min(max(level, 1), 6)
            # 计算原行中的起止列
            marker_start_col = len(line) - len(stripped)
            marker_end_col = marker_start_col + level
            # 文本起点：跳过可选空白
            text_start_col = marker_end_col
            while text_start_col < len(line) and line[text_start_col] in (" ", "\t", "\u00a0", "\u3000"):
                text_start_col += 1
            text_end_col = len(line.rstrip('\n'))
            if text_start_col < text_end_col:
                self._text.tag_add(f'heading{level}_marker',
                                   f"{line_num}.{marker_start_col}",
                                   f"{line_num}.{marker_end_col}")
                self._text.tag_add(f'heading{level}',
                                   f"{line_num}.{text_start_col}",
                                   f"{line_num}.{text_end_col}")
                return
        
        # 分隔线
        if re.match(PATTERNS['hr'], line):
            self._text.tag_add('hr', line_start, f"{line_num}.end")
            return
        
        # 引用
        match = re.match(PATTERNS['blockquote'], line)
        if match:
            marker_end = f"{line_num}.{len(match.group(0))}"
            self._text.tag_add('blockquote_marker', line_start, marker_end)
            self._text.tag_add('blockquote', line_start, f"{line_num}.end")
        
        # 任务列表
        match = re.match(PATTERNS['task_list'], line)
        if match:
            indent_len = len(match.group(1))
            marker_start = f"{line_num}.{indent_len}"
            checkbox_end = f"{line_num}.{match.end()}"
            self._text.tag_add('list_marker', marker_start, f"{line_num}.{indent_len + 1}")
            self._text.tag_add('task_checkbox', f"{line_num}.{indent_len + 2}", checkbox_end)
            if match.group(3).lower() == 'x':
                self._text.tag_add('task_done', checkbox_end, f"{line_num}.end")
            return
        
        # 无序列表
        match = re.match(PATTERNS['unordered_list'], line)
        if match:
            indent_len = len(match.group(1))
            marker_start = f"{line_num}.{indent_len}"
            marker_end = f"{line_num}.{indent_len + 1}"
            self._text.tag_add('list_marker', marker_start, marker_end)
        
        # 有序列表
        match = re.match(PATTERNS['ordered_list'], line)
        if match:
            indent_len = len(match.group(1))
            marker_start = f"{line_num}.{indent_len}"
            marker_end = f"{line_num}.{indent_len + len(match.group(2))}"
            self._text.tag_add('list_marker', marker_start, marker_end)
        
        # 表格分隔符
        if re.match(PATTERNS['table_separator'], line):
            self._text.tag_add('table_separator', line_start, f"{line_num}.end")
            return
        
        # 表格行
        if re.match(PATTERNS['table_row'], line):
            self._text.tag_add('table', line_start, f"{line_num}.end")
        
        # 行内元素
        self._highlight_inline(line, line_num)
    
    def _highlight_inline(self, line: str, line_num: int):
        """高亮行内元素"""
        # 图片 (必须在链接之前)
        for match in re.finditer(PATTERNS['image'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('image', start, end)
        
        # 链接
        for match in re.finditer(PATTERNS['link'], line):
            # 检查是否是图片的一部分
            if match.start() > 0 and line[match.start() - 1] == '!':
                continue
            
            full_start = f"{line_num}.{match.start()}"
            text_end = f"{line_num}.{match.start() + 1 + len(match.group(1))}"
            url_start = f"{line_num}.{match.start() + 2 + len(match.group(1))}"
            full_end = f"{line_num}.{match.end()}"
            
            self._text.tag_add('link_bracket', full_start, f"{line_num}.{match.start() + 1}")
            self._text.tag_add('link_text', f"{line_num}.{match.start() + 1}", text_end)
            self._text.tag_add('link_bracket', text_end, url_start)
            self._text.tag_add('link_url', url_start, f"{line_num}.{match.end() - 1}")
            self._text.tag_add('link_bracket', f"{line_num}.{match.end() - 1}", full_end)
        
        # 数学公式 (块级)
        for match in re.finditer(PATTERNS['math_block'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('math', start, end)
        
        # 数学公式 (行内)
        for match in re.finditer(PATTERNS['math_inline'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('math', start, end)
        
        # 行内代码
        for match in re.finditer(PATTERNS['code_inline'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('code_inline', start, end)
        
        # 粗斜体 (必须在粗体和斜体之前)
        for match in re.finditer(PATTERNS['bold_italic'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('bold_italic', start, end)
        
        # 粗体
        for match in re.finditer(PATTERNS['bold'], line):
            # 检查是否是粗斜体的一部分
            if match.start() > 0 and line[match.start() - 1] == '*':
                continue
            if match.end() < len(line) and line[match.end()] == '*':
                continue
            
            start = f"{line_num}.{match.start()}"
            marker1_end = f"{line_num}.{match.start() + 2}"
            text_end = f"{line_num}.{match.end() - 2}"
            end = f"{line_num}.{match.end()}"
            
            self._text.tag_add('bold_marker', start, marker1_end)
            self._text.tag_add('bold', marker1_end, text_end)
            self._text.tag_add('bold_marker', text_end, end)
        
        # 斜体
        for match in re.finditer(PATTERNS['italic'], line):
            # 检查是否是粗体或粗斜体的一部分
            if match.start() > 0 and line[match.start() - 1] == '*':
                continue
            if match.end() < len(line) and line[match.end()] == '*':
                continue
            
            start = f"{line_num}.{match.start()}"
            marker1_end = f"{line_num}.{match.start() + 1}"
            text_end = f"{line_num}.{match.end() - 1}"
            end = f"{line_num}.{match.end()}"
            
            self._text.tag_add('italic_marker', start, marker1_end)
            self._text.tag_add('italic', marker1_end, text_end)
            self._text.tag_add('italic_marker', text_end, end)
        
        # 删除线
        for match in re.finditer(PATTERNS['strikethrough'], line):
            start = f"{line_num}.{match.start()}"
            end = f"{line_num}.{match.end()}"
            self._text.tag_add('strikethrough', start, end)
    
    def set_theme(self, theme: HighlightTheme):
        """设置主题"""
        self.theme = theme
        self._configure_tags()
        self.highlight_all()
    
    def enable(self):
        """启用语法高亮"""
        self._enabled = True
        self.highlight_all()
    
    def disable(self):
        """禁用语法高亮"""
        self._enabled = False
        # 清除所有标签
        for tag in self._text.tag_names():
            if tag != 'sel':
                self._text.tag_remove(tag, "1.0", "end")


class LineNumbers:
    """行号显示组件"""
    
    def __init__(self, text_widget, parent=None):
        """
        初始化行号组件
        
        Args:
            text_widget: 关联的文本组件
            parent: 父容器
        """
        self.text_widget = text_widget
        self._enabled = True
        
        # 获取底层 Text 组件
        if hasattr(text_widget, '_textbox'):
            self._text = text_widget._textbox
        else:
            self._text = text_widget
        
        # 创建行号画布
        self.canvas = tk.Canvas(
            parent or self._text.master,
            width=50,
            bg='#f9fafb',
            highlightthickness=0
        )
        
        # 绑定事件
        self._text.bind('<Configure>', self._on_configure)
        self._text.bind('<KeyRelease>', self._update)
        self._text.bind('<<Modified>>', self._on_modified)
        
        # 绑定滚动
        self._text.bind('<MouseWheel>', self._on_scroll)
        self._text.bind('<Button-4>', self._on_scroll)
        self._text.bind('<Button-5>', self._on_scroll)
    
    def _on_configure(self, event=None):
        """配置变化事件"""
        self._update()
    
    def _on_modified(self, event=None):
        """文本修改事件"""
        if self._text.edit_modified():
            self._text.edit_modified(False)
            self._update()
    
    def _on_scroll(self, event=None):
        """滚动事件"""
        self._text.after(10, self._update)
    
    def _update(self, event=None):
        """更新行号显示"""
        if not self._enabled:
            return
        
        self.canvas.delete("all")
        
        # 获取可见区域
        first_visible = self._text.index("@0,0")
        first_line = int(first_visible.split('.')[0])
        
        # 获取总行数
        total_lines = int(self._text.index('end-1c').split('.')[0])
        
        # 计算行高
        try:
            bbox = self._text.bbox("1.0")
            if bbox:
                line_height = bbox[3]
            else:
                line_height = 20
        except:
            line_height = 20
        
        # 绘制行号
        y = 0
        for line_num in range(first_line, total_lines + 1):
            try:
                bbox = self._text.bbox(f"{line_num}.0")
                if bbox:
                    y = bbox[1]
                    self.canvas.create_text(
                        45, y + line_height // 2,
                        text=str(line_num),
                        anchor='e',
                        fill='#9ca3af',
                        font=('Consolas', 10)
                    )
                else:
                    break
            except:
                break
    
    def show(self):
        """显示行号"""
        self._enabled = True
        self.canvas.pack(side='left', fill='y', before=self.text_widget)
        self._update()
    
    def hide(self):
        """隐藏行号"""
        self._enabled = False
        self.canvas.pack_forget()
    
    def toggle(self):
        """切换显示状态"""
        if self._enabled:
            self.hide()
        else:
            self.show()
