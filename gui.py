# -*- coding: utf-8 -*-

import os
import re
import sys
import json
import threading
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk
import markdown as md_parser
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')  # 抑制所有警告
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = ['SimHei', 'DejaVu Sans']  # 支持中文
import logging
logging.getLogger('matplotlib').setLevel(logging.ERROR)  # 抑制matplotlib日志

# 导入转换器模块
from converter import MarkdownToWordConverter
from parser import parse_markdown, parse_inline, parse_table, InlineType
from utils import normalize_markdown, convert_latex_delimiters

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.md2word_config.json')

# ============== 主题配置 ==============
ctk.set_appearance_mode("light")  # 亮色主题
ctk.set_default_color_theme("blue")

# 亮色主题颜色
COLORS_LIGHT = {
    'primary': '#6366F1',       # 靖蓝紫 - 主色
    'primary_hover': '#4F46E5',
    'secondary': '#EC4899',      # 粉色 - 强调色
    'success': '#10B981',        # 翠绿
    'warning': '#F59E0B',        # 琥珀
    'danger': '#EF4444',         # 红色
    'bg_light': '#F8FAFC',       # 浅灰背景
    'bg_card': '#FFFFFF',        # 卡片白
    'bg_sidebar': '#F1F5F9',     # 侧边栏背景
    'text_primary': '#1E293B',   # 深灰文字
    'text_secondary': '#64748B', # 次要文字
    'text_muted': '#94A3B8',     # 更淡的文字
    'border': '#E2E8F0',         # 边框
    'border_focus': '#6366F1',   # 聚焦边框
    'line_number': '#94A3B8',    # 行号颜色
    'line_number_bg': '#F1F5F9', # 行号背景
    'highlight': '#FEF3C7',      # 高亮背景
    'shadow': '#E5E7EB',         # 阴影
    'editor_bg': '#F8FAFC',      # 编辑器背景
    'preview_bg': '#FFFFFF',     # 预览区背景
}

# 暗色主题颜色
COLORS_DARK = {
    'primary': '#818CF8',        # 浅蓝紫 - 主色
    'primary_hover': '#6366F1',
    'secondary': '#F472B6',      # 浅粉色 - 强调色
    'success': '#34D399',        # 翠绿
    'warning': '#FBBF24',        # 琥珀
    'danger': '#F87171',         # 红色
    'bg_light': '#1E293B',       # 深色背景
    'bg_card': '#334155',        # 卡片深色
    'bg_sidebar': '#1E293B',     # 侧边栏背景
    'text_primary': '#F1F5F9',   # 浅色文字
    'text_secondary': '#CBD5E1', # 次要文字
    'text_muted': '#94A3B8',     # 更淡的文字
    'border': '#475569',         # 边框
    'border_focus': '#818CF8',   # 聚焦边框
    'line_number': '#94A3B8',    # 行号颜色
    'line_number_bg': '#1E293B', # 行号背景
    'highlight': '#854D0E',      # 高亮背景
    'shadow': '#0F172A',         # 阴影
    'editor_bg': '#1E293B',      # 编辑器背景
    'preview_bg': '#334155',     # 预览区背景
}

# 当前使用的颜色配置（默认亮色）
COLORS = COLORS_LIGHT.copy()


def load_config() -> dict:
    """加载配置文件"""
    default_config = {
        'recent_files': [],
        'font_size': 14,
        'theme': 'light',
        'sidebar_visible': True,
        'sidebar_width': 250,
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 合并默认配置
                return {**default_config, **config}
    except (json.JSONDecodeError, IOError, OSError):
        pass
    return default_config


def save_config(config: dict):
    """保存配置文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except (IOError, OSError):
        pass


class ModernButton(ctk.CTkButton):
    """现代化按钮组件"""
    def __init__(self, master, text, command=None, style="primary", icon=None, **kwargs):
        colors = {
            'primary': (COLORS['primary'], COLORS['primary_hover']),
            'secondary': (COLORS['secondary'], '#DB2777'),
            'success': (COLORS['success'], '#059669'),
            'outline': ('transparent', COLORS['bg_light']),
        }
        
        fg_color, hover_color = colors.get(style, colors['primary'])
        text_color = 'white' if style != 'outline' else COLORS['primary']
        border_width = 2 if style == 'outline' else 0
        
        super().__init__(
            master,
            text=text,
            command=command,
            fg_color=fg_color,
            hover_color=hover_color,
            text_color=text_color,
            border_width=border_width,
            border_color=COLORS['primary'] if style == 'outline' else None,
            corner_radius=12,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            **kwargs
        )


class ModernCard(ctk.CTkFrame):
    """现代化卡片容器"""
    def __init__(self, master, title=None, **kwargs):
        super().__init__(
            master,
            fg_color=COLORS['bg_card'],
            corner_radius=16,
            border_width=1,
            border_color=COLORS['border'],
            **kwargs
        )
        
        if title:
            self.title_label = ctk.CTkLabel(
                self,
                text=title,
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color=COLORS['text_primary']
            )
            self.title_label.pack(anchor="w", padx=20, pady=(15, 10))


class LineNumberedText(ctk.CTkFrame):
    """带行号的文本编辑器 - 精确对齐版"""
    def __init__(self, master, font_size=14, on_scroll=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self.font_size = font_size
        self.on_scroll_callback = on_scroll  # 滚动回调
        
        # 使用原生 tk.Text 而不是 CTkTextbox，以便精确控制
        # 容器框架
        self.container = tk.Frame(self, bg=COLORS['bg_light'])
        self.container.pack(fill='both', expand=True)
        
        # 行号栏
        self.line_numbers = tk.Text(
            self.container,
            width=4,
            padx=4,
            pady=5,
            takefocus=0,
            border=0,
            background=COLORS['line_number_bg'],
            foreground=COLORS['line_number'],
            state='disabled',
            wrap='none',
            font=('Consolas', font_size),
            cursor='arrow',
        )
        self.line_numbers.pack(side='left', fill='y')
        
        # 主文本区 - 使用原生 tk.Text
        self.text_frame = tk.Frame(self.container, bg=COLORS['bg_light'])
        self.text_frame.pack(side='left', fill='both', expand=True)
        
        self._textbox = tk.Text(
            self.text_frame,
            font=('Consolas', font_size),
            bg=COLORS['bg_light'],
            fg=COLORS['text_primary'],
            wrap='word',
            border=0,
            padx=8,
            pady=5,
            undo=True,
            autoseparators=True,  # 自动分隔撤销点
            maxundo=-1,  # 无限撤销
            insertbackground=COLORS['text_primary'],
        )
        
        # 滚动条
        self.scrollbar = tk.Scrollbar(self.text_frame, command=self._on_scrollbar)
        self.scrollbar.pack(side='right', fill='y')
        self._textbox.pack(side='left', fill='both', expand=True)
        self._textbox.config(yscrollcommand=self._on_text_scroll)
        
        # 绑定事件
        self._textbox.bind('<KeyRelease>', self._on_change)
        self._textbox.bind('<MouseWheel>', self._on_mousewheel)
        self._textbox.bind('<Configure>', self._on_change)
        self.line_numbers.bind('<MouseWheel>', self._on_mousewheel)
        
        # 初始化行号
        self.after(50, self._update_line_numbers)
    
    def _on_scrollbar(self, *args):
        """滚动条操作"""
        self._textbox.yview(*args)
        self.line_numbers.yview(*args)
    
    def _on_text_scroll(self, first, last):
        """文本滚动同步"""
        self.scrollbar.set(first, last)
        self.line_numbers.yview_moveto(first)
        # 通知外部滚动回调
        if self.on_scroll_callback:
            self.on_scroll_callback(float(first))
    
    def _on_mousewheel(self, event):
        """鼠标滚轮同步"""
        self._textbox.yview_scroll(-1 * (event.delta // 120), "units")
        self.line_numbers.yview_scroll(-1 * (event.delta // 120), "units")
        # 通知外部滚动回调
        if self.on_scroll_callback:
            first = self._textbox.yview()[0]
            self.on_scroll_callback(first)
        return "break"
    
    def _on_change(self, event=None):
        """内容变化时更新行号"""
        self.after(5, self._update_line_numbers)
    
    def _update_line_numbers(self):
        """更新行号显示"""
        self.line_numbers.config(state='normal')
        self.line_numbers.delete('1.0', 'end')
        
        # 获取文本行数
        content = self._textbox.get('1.0', 'end-1c')
        line_count = content.count('\n') + 1
        
        # 生成行号
        line_numbers_str = '\n'.join(str(i) for i in range(1, line_count + 1))
        self.line_numbers.insert('1.0', line_numbers_str)
        self.line_numbers.config(state='disabled')
        
        # 同步滚动位置
        self.line_numbers.yview_moveto(self._textbox.yview()[0])
    
    def get(self, start, end):
        """获取文本"""
        return self._textbox.get(start, end)
    
    def insert(self, index, text):
        """插入文本"""
        self._textbox.insert(index, text)
        self._update_line_numbers()
    
    def delete(self, start, end):
        """删除文本"""
        self._textbox.delete(start, end)
        self._update_line_numbers()
    
    def bind(self, event, callback):
        """绑定事件"""
        self._textbox.bind(event, callback)
    
    def set_font_size(self, size):
        """设置字体大小"""
        self.font_size = size
        self._textbox.configure(font=('Consolas', size))
        self.line_numbers.configure(font=('Consolas', size))
        self._update_line_numbers()


class MarkdownPreview(ctk.CTkFrame):
    """可编辑的Markdown预览组件 - 支持双向同步"""
    def __init__(self, master, on_content_change=None, **kwargs):
        super().__init__(master, fg_color=COLORS['bg_card'], corner_radius=12, **kwargs)
        
        # 内容变化回调
        self.on_content_change = on_content_change
        
        # 是否正在更新（防止循环触发）
        self._updating = False
        
        # 存储段落信息：{line_start: {'type': 'paragraph', 'md_line': 1, 'format': []}}
        self.paragraph_map = {}
        
        # 使用 Text widget 支持富文本，整体模拟排版后的页面效果
        self.text = tk.Text(
            self,
            wrap='word',
            bg='#FFFFFF',
            fg='#111827',
            font=('宋体', 16),
            width=60,
            padx=20,
            pady=30,
            relief='flat',
            cursor='xterm',  # 可编辑光标
            spacing1=0,
            spacing3=0,
            undo=True,  # 启用撤销
        )
        
        # 滚动条
        self.scrollbar = ctk.CTkScrollbar(self, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        self.text.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        
        # 配置文本标签样式 - 与Word导出尽量保持一致
        self._setup_tags()
        
        # 存储公式图片引用（防止被垃圾回收）
        self.math_images = []
        
        # 公式计数器
        self.equation_counter = 0
        
        # 绑定编辑事件
        self.text.bind('<KeyRelease>', self._on_text_change)
        self.text.bind('<ButtonRelease-1>', self._on_text_change)
        
        # 右键菜单
        self._create_context_menu()
    
    def _create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="加粗 (Ctrl+B)", command=self._toggle_bold)
        self.context_menu.add_command(label="斜体 (Ctrl+I)", command=self._toggle_italic)
        self.context_menu.add_command(label="删除线", command=self._toggle_strikethrough)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="上标 X²", command=self._toggle_superscript)
        self.context_menu.add_command(label="下标 X₂", command=self._toggle_subscript)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="复制", command=lambda: self.text.event_generate('<<Copy>>'))
        self.context_menu.add_command(label="粘贴", command=lambda: self.text.event_generate('<<Paste>>'))
        
        self.text.bind('<Button-3>', self._show_context_menu)
        self.text.bind('<Control-b>', lambda e: self._toggle_bold())
        self.text.bind('<Control-i>', lambda e: self._toggle_italic())
    
    def _show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def _toggle_bold(self):
        """切换加粗"""
        try:
            sel_start = self.text.index(tk.SEL_FIRST)
            sel_end = self.text.index(tk.SEL_LAST)
            tags = self.text.tag_names(sel_start)
            
            if 'bold' in tags:
                self.text.tag_remove('bold', sel_start, sel_end)
            else:
                self.text.tag_add('bold', sel_start, sel_end)
            
            self._sync_to_markdown()
        except tk.TclError:
            pass  # 没有选中文本
    
    def _toggle_italic(self):
        """切换斜体"""
        try:
            sel_start = self.text.index(tk.SEL_FIRST)
            sel_end = self.text.index(tk.SEL_LAST)
            tags = self.text.tag_names(sel_start)
            
            if 'italic' in tags:
                self.text.tag_remove('italic', sel_start, sel_end)
            else:
                self.text.tag_add('italic', sel_start, sel_end)
            
            self._sync_to_markdown()
        except tk.TclError:
            pass
    
    def _toggle_strikethrough(self):
        """切换删除线"""
        try:
            sel_start = self.text.index(tk.SEL_FIRST)
            sel_end = self.text.index(tk.SEL_LAST)
            tags = self.text.tag_names(sel_start)
            
            if 'strikethrough' in tags:
                self.text.tag_remove('strikethrough', sel_start, sel_end)
            else:
                self.text.tag_add('strikethrough', sel_start, sel_end)
            
            self._sync_to_markdown()
        except tk.TclError:
            pass
    
    def _toggle_superscript(self):
        """切换上标"""
        try:
            sel_start = self.text.index(tk.SEL_FIRST)
            sel_end = self.text.index(tk.SEL_LAST)
            tags = self.text.tag_names(sel_start)
            
            if 'superscript' in tags:
                self.text.tag_remove('superscript', sel_start, sel_end)
            else:
                # 移除下标（互斥）
                self.text.tag_remove('subscript', sel_start, sel_end)
                self.text.tag_add('superscript', sel_start, sel_end)
            
            self._sync_to_markdown()
        except tk.TclError:
            pass
    
    def _toggle_subscript(self):
        """切换下标"""
        try:
            sel_start = self.text.index(tk.SEL_FIRST)
            sel_end = self.text.index(tk.SEL_LAST)
            tags = self.text.tag_names(sel_start)
            
            if 'subscript' in tags:
                self.text.tag_remove('subscript', sel_start, sel_end)
            else:
                # 移除上标（互斥）
                self.text.tag_remove('superscript', sel_start, sel_end)
                self.text.tag_add('subscript', sel_start, sel_end)
            
            self._sync_to_markdown()
        except tk.TclError:
            pass
    
    def _on_text_change(self, event=None):
        """文本变化时触发同步"""
        if self._updating:
            return
        # 用防抖延迟同步
        if hasattr(self, '_sync_timer'):
            self.after_cancel(self._sync_timer)
        self._sync_timer = self.after(500, self._sync_to_markdown)
    
    def _sync_to_markdown(self):
        """将预览区内容同步回Markdown"""
        if self._updating or not self.on_content_change:
            return
        
        try:
            # 获取所有文本并转换为Markdown
            markdown_text = self._convert_to_markdown()
            if markdown_text:
                self.on_content_change(markdown_text)
        except Exception:
            pass
    
    def _convert_to_markdown(self) -> str:
        """将富文本转换为Markdown"""
        result = []
        
        # 获取所有文本
        content = self.text.get("1.0", "end-1c")
        lines = content.split('\n')
        
        for line_num, line in enumerate(lines, 1):
            if not line.strip():
                result.append('')
                continue
            
            # 检查这行的标签
            line_start = f"{line_num}.0"
            line_tags = self.text.tag_names(line_start)
            
            # 检查是否是标题
            if 'h1' in line_tags:
                result.append(f"# {line}")
            elif 'h2' in line_tags:
                result.append(f"## {line}")
            elif 'h3' in line_tags:
                result.append(f"### {line}")
            elif 'h4' in line_tags:
                result.append(f"#### {line}")
            else:
                # 处理行内格式
                formatted_line = self._format_line(line_num, line)
                result.append(formatted_line)
        
        return '\n'.join(result)
    
    def _format_line(self, line_num: int, line: str) -> str:
        """处理行内格式（粗体、斜体、上下标等）"""
        if not line:
            return line
        
        # 获取这行每个字符的标签
        segments = []  # [(text, tags), ...]
        current_text = ""
        current_tags = set()
        
        for col in range(len(line)):
            pos = f"{line_num}.{col}"
            char_tags = set(self.text.tag_names(pos))
            # 只关注格式标签（添加上下标）
            format_tags = char_tags & {'bold', 'italic', 'strikethrough', 'code', 'superscript', 'subscript'}
            
            if format_tags != current_tags:
                if current_text:
                    segments.append((current_text, current_tags))
                current_text = line[col]
                current_tags = format_tags
            else:
                current_text += line[col]
        
        if current_text:
            segments.append((current_text, current_tags))
        
        # 将段落转换为Markdown/HTML
        result = ""
        for text, tags in segments:
            formatted = text
            if 'bold' in tags and 'italic' in tags:
                formatted = f"***{text}***"
            elif 'bold' in tags:
                formatted = f"**{text}**"
            elif 'italic' in tags:
                formatted = f"*{text}*"
            if 'strikethrough' in tags:
                formatted = f"~~{formatted}~~"
            if 'code' in tags:
                formatted = f"`{text}`"
            if 'superscript' in tags:
                formatted = f"<sup>{text}</sup>"
            if 'subscript' in tags:
                formatted = f"<sub>{text}</sub>"
            result += formatted
        
        return result
    
    def set_updating(self, updating: bool):
        """设置更新状态（防止循环触发）"""
        self._updating = updating
    
    def _setup_tags(self):
        """配置文本标签样式 - 模拟Word中的样式"""
        # 标题样式 - 相比正文放大一到两个级别，比例更接近论文
        self.text.tag_configure('h1', font=('黑体', 28, 'bold'), justify='center', spacing1=28, spacing3=16)
        self.text.tag_configure('h2', font=('黑体', 22, 'bold'), justify='center', spacing1=22, spacing3=14)
        self.text.tag_configure('h3', font=('黑体', 18, 'bold'), spacing1=16, spacing3=12)
        self.text.tag_configure('h4', font=('黑体', 16, 'bold'), spacing1=14, spacing3=10)
        
        # 正文样式：统一控制左右页边距和段前段后间距，模拟 LaTeX/Word 正文
        # 首行缩进2字符（16pt字号 × 2 ≈ 32像素）
        self.text.tag_configure(
            'body',
            font=('宋体', 16),
            lmargin1=112,  # 左边距（首行）= 80 + 32（首行缩进）
            lmargin2=80,   # 左边距（后续行）
            rmargin=80,    # 右边距
            spacing1=4,    # 段前
            spacing3=4,    # 段后
        )
        
        # 粗体、斜体（保持与正文字号一致）
        self.text.tag_configure('bold', font=('宋体', 16, 'bold'))
        self.text.tag_configure('italic', font=('宋体', 16, 'italic'))
        self.text.tag_configure('bold_italic', font=('宋体', 16, 'bold italic'))
        
        # 代码（白色底色）
        self.text.tag_configure('code', font=('Consolas', 10), background='#F5F5F5')
        self.text.tag_configure('code_block', font=('Consolas', 10), background='#FAFAFA', foreground='#1F2937')
        
        # 公式：优先使用 Cambria Math，回退到其他数学字体
        math_font = self._get_math_font()
        self.text.tag_configure('math', font=(math_font, 16), foreground='#1a1a2e')
        self.text.tag_configure('math_block', font=(math_font, 18), foreground='#1a1a2e', justify='center', spacing1=8, spacing3=8)
        
        # 链接
        self.text.tag_configure('link', foreground='#0000FF', underline=True)
        
        # 删除线
        self.text.tag_configure('strikethrough', overstrike=True)
        
        # 上标和下标
        self.text.tag_configure('superscript', font=('宋体', 9), offset=6)  # 上标：更小字体，向上偏移
        self.text.tag_configure('subscript', font=('宋体', 9), offset=-3)  # 下标：更小字体，向下偏移
        
        # 引用
        self.text.tag_configure('quote', font=('宋体', 11, 'italic'), foreground='#6B7280', lmargin1=30, lmargin2=30)
        
        # 列表：在正文基础上增加缩进
        self.text.tag_configure(
            'list_item',
            font=('宋体', 16),
            lmargin1=36,
            lmargin2=60,
            spacing1=2,
            spacing3=2,
        )

        # 表格整体居中显示
        self.text.tag_configure('table_block', justify='center')
        
        # 提高上下标标签的优先级，确保字体大小生效
        self.text.tag_raise('superscript')
        self.text.tag_raise('subscript')
    
    def _get_math_font(self) -> str:
        """获取可用的数学字体，按优先级尝试"""
        import tkinter.font as tkfont
        available = set(tkfont.families())
        # 按优先级尝试数学字体
        for font in ['Cambria Math', 'STIX Two Math', 'Latin Modern Math', 'Times New Roman', 'SimSun']:
            if font in available:
                return font
        return 'TkDefaultFont'
    
    def update_preview(self, markdown_text: str):
        """更新预览内容 - 使用共用解析器渲染"""
        # 允许选中复制，因此不再切换到 disabled，仅在这里完全重绘
        self.text.delete('1.0', 'end')
        
        # 清除旧的公式图片，重置计数器
        self.math_images = []
        self.equation_counter = 0
        
        # 预处理文本
        markdown_text = convert_latex_delimiters(markdown_text)  # 转换 \(...\) 和 \[...\] 格式
        markdown_text = normalize_markdown(markdown_text)  # 规范化格式
        
        # 使用共用解析器解析
        blocks = parse_markdown(markdown_text)
        
        for block in blocks:
            self._render_block(block)
    
    def _render_block(self, block):
        """渲染块级元素"""
        if block.type == 'heading':
            self._insert_heading(block.content, block.level)
        
        elif block.type == 'paragraph':
            self._insert_paragraph(block.content)
        
        elif block.type == 'code_block':
            self._insert_code_block(block.content, block.language)
        
        elif block.type == 'math_block':
            self._insert_math_block(block.content)
        
        elif block.type == 'table':
            self._insert_table(block.content)
        
        elif block.type == 'quote':
            self._insert_quote(block.content)
        
        elif block.type == 'list':
            # level=0 无序列表，level=1 有序列表
            self._insert_list(block.content, ordered=(block.level == 1))
        
        elif block.type == 'image':
            self._insert_image(block.content, block.language)  # language存储了url
        
        elif block.type == 'hr':
            self.text.insert('end', '\n' + '─' * 50 + '\n\n')
    
    def _insert_heading(self, text: str, level: int):
        """插入标题 - 使用共用解析器处理行内元素"""
        tag = f'h{min(level, 4)}'
        self._insert_inline_elements(text, extra_tags=[tag])
        self.text.insert('end', '\n\n')
    
    def _insert_paragraph(self, text: str):
        """插入段落 - 使用共用解析器处理行内元素"""
        # 使用 body 标签统一控制段落字体、边距和行距
        self._insert_inline_elements(text, extra_tags=['body'])
        self.text.insert('end', '\n')
    
    def _insert_inline_elements(self, text: str, extra_tags: list = None):
        """插入行内元素 - 使用共用解析器，与Word导出逻辑一致"""
        elements = parse_inline(text)
        
        for elem in elements:
            tags = list(extra_tags) if extra_tags else []
            
            if elem.type == InlineType.TEXT:
                self.text.insert('end', elem.content, tuple(tags) if tags else ())
            
            elif elem.type == InlineType.BOLD:
                self.text.insert('end', elem.content, tuple(tags + ['bold']))
            
            elif elem.type == InlineType.ITALIC:
                self.text.insert('end', elem.content, tuple(tags + ['italic']))
            
            elif elem.type == InlineType.BOLD_ITALIC:
                self.text.insert('end', elem.content, tuple(tags + ['bold_italic']))
            
            elif elem.type == InlineType.CODE:
                self.text.insert('end', elem.content, tuple(tags + ['code']))
            
            elif elem.type == InlineType.MATH:
                # 行内公式：优先用 matplotlib 渲染为图片，失败则退回数学字体文本
                formula_text = elem.content.strip()
                img = self._render_latex(formula_text, is_inline=True)
                if img:
                    self.math_images.append(img)
                    self.text.image_create('end', image=img)
                else:
                    self.text.insert('end', formula_text, tuple(tags + ['math']))
            
            elif elem.type == InlineType.LINK:
                self.text.insert('end', elem.content, tuple(tags + ['link']))
            
            elif elem.type == InlineType.IMAGE:
                self.text.insert('end', f'🖼️[{elem.content}]', tuple(tags))
            
            elif elem.type == InlineType.STRIKETHROUGH:
                self.text.insert('end', elem.content, tuple(tags + ['strikethrough']))
            
            elif elem.type == InlineType.SUPERSCRIPT:
                self.text.insert('end', elem.content, tuple(tags + ['superscript']))
            
            elif elem.type == InlineType.SUBSCRIPT:
                self.text.insert('end', elem.content, tuple(tags + ['subscript']))
    
    def _insert_code_block(self, code: str, language: str = ''):
        """插入代码块"""
        if language:
            self.text.insert('end', f'[{language}]\n', ('code',))
        self.text.insert('end', code + '\n\n', ('code_block',))
    
    def _render_latex(self, latex: str, fontsize: int = 16, is_inline: bool = False) -> ImageTk.PhotoImage:
        """使用 matplotlib.mathtext 渲染 LaTeX 公式为图片。

        渲染失败时返回 None，由调用方回退为纯文本显示。
        """
        try:
            # 使用 Computer Modern 数学字体族，效果接近 LaTeX/Word 公式
            plt.rcParams['mathtext.fontset'] = 'cm'
            plt.rcParams['mathtext.rm'] = 'serif'
            plt.rcParams['mathtext.it'] = 'serif:italic'
            plt.rcParams['mathtext.bf'] = 'serif:bold'

            fig, ax = plt.subplots(figsize=(0.01, 0.01))
            ax.axis('off')

            # 行内公式略小，块级公式略大，以匹配正文 16pt 的视觉大小
            render_fontsize = 14 if is_inline else 18
            formula = latex.strip()

            text = ax.text(
                0.5,
                0.5,
                f'${formula}$',
                fontsize=render_fontsize,
                ha='center',
                va='center',
                transform=ax.transAxes,
                color='#1a1a2e',
            )

            fig.canvas.draw()
            bbox = text.get_window_extent(fig.canvas.get_renderer())
            bbox = bbox.expanded(1.05, 1.10)
            fig.set_size_inches(bbox.width / fig.dpi, bbox.height / fig.dpi)

            buf = BytesIO()
            fig.savefig(
                buf,
                format='png',
                bbox_inches='tight',
                pad_inches=0.02,
                dpi=120,
                transparent=True,
            )
            plt.close(fig)

            buf.seek(0)
            img = Image.open(buf)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            # 失败时在控制台输出原因，预览中由调用方回退为文本
            print(f"[math render error] {e}")
            return None

    def _insert_math_block(self, formula: str):
        """插入块级公式：居中显示，右侧编号"""
        self.equation_counter += 1
        formula_text = formula.strip()
        
        if not formula_text:
            return
        
        # 尝试渲染简单公式为图片
        img = self._render_latex(formula_text, is_inline=False)
        
        self.text.insert('end', '\n')
        
        if img:
            # 图片渲染成功
            self.math_images.append(img)
            self.text.insert('end', '          ')
            self.text.image_create('end', image=img)
            self.text.insert('end', f'    ({self.equation_counter})\n')
        else:
            # 直接插入文本，使用 math_block 标签居中
            display = f'    {formula_text}    ({self.equation_counter})'
            self.text.insert('end', display, ('math_block',))
            self.text.insert('end', '\n')
        
        self.text.insert('end', '\n')
    
    def _insert_table(self, table_text: str):
        """插入表格 - 在预览中渲染为网格表，支持格式化内容"""
        headers, rows, alignments = parse_table(table_text)
        if not headers:
            self.text.insert('end', table_text + '\n\n', ('body',))
            return

        # 创建容器Frame用于居中
        container = tk.Frame(self.text, bg=self.text.cget('bg'))
        
        # 创建表格 Frame
        table_frame = tk.Frame(container, bg='#FFFFFF', bd=0)
        table_frame.pack(anchor='center')

        all_rows = [headers] + rows
        num_cols = len(headers)

        for r, row in enumerate(all_rows):
            for c in range(num_cols):
                cell_text = row[c] if c < len(row) else ''
                is_header = (r == 0)
                
                # 使用 Text 组件来支持格式化内容
                cell = tk.Text(
                    table_frame,
                    font=('黑体' if is_header else '宋体', 14),
                    fg='#1E293B',
                    bg='#F1F5F9' if is_header else '#FFFFFF',
                    bd=1,
                    relief='solid',
                    padx=10,
                    pady=4,
                    width=12,
                    height=1,
                    wrap='none',
                    cursor='arrow',
                )
                
                # 配置标签样式
                cell.tag_configure('bold', font=('黑体' if is_header else '宋体', 14, 'bold'))
                cell.tag_configure('italic', font=('黑体' if is_header else '宋体', 14, 'italic'))
                cell.tag_configure('bold_italic', font=('黑体' if is_header else '宋体', 14, 'bold italic'))
                cell.tag_configure('superscript', font=('宋体', 10), offset=4)
                cell.tag_configure('subscript', font=('宋体', 10), offset=-3)
                cell.tag_configure('strikethrough', overstrike=True)
                cell.tag_configure('code', font=('Consolas', 12), background='#F0F0F0')
                
                # 解析并插入格式化内容
                self._insert_cell_content(cell, cell_text, is_header)
                
                cell.config(state='disabled')  # 禁用编辑
                cell.grid(row=r, column=c, sticky='nsew')

        # 插入到 Text 中居中显示
        self.text.insert('end', '\n')
        self.text.window_create('end', window=container)
        self.text.insert('end', '\n\n')
        
        # 设置容器宽度以实现居中（延迟执行以获取实际宽度）
        def center_table():
            try:
                text_width = self.text.winfo_width()
                if text_width > 100:
                    container.configure(width=text_width - 40)
            except:
                pass
        self.text.after(10, center_table)
    
    def _insert_cell_content(self, cell: tk.Text, text: str, is_header: bool = False):
        """在表格单元格中插入格式化内容"""
        from parser import parse_inline, InlineType
        
        elements = parse_inline(text)
        
        for elem in elements:
            if elem.type == InlineType.TEXT:
                cell.insert('end', elem.content)
            elif elem.type == InlineType.BOLD:
                cell.insert('end', elem.content, ('bold',))
            elif elem.type == InlineType.ITALIC:
                cell.insert('end', elem.content, ('italic',))
            elif elem.type == InlineType.BOLD_ITALIC:
                cell.insert('end', elem.content, ('bold_italic',))
            elif elem.type == InlineType.SUPERSCRIPT:
                cell.insert('end', elem.content, ('superscript',))
            elif elem.type == InlineType.SUBSCRIPT:
                cell.insert('end', elem.content, ('subscript',))
            elif elem.type == InlineType.STRIKETHROUGH:
                cell.insert('end', elem.content, ('strikethrough',))
            elif elem.type == InlineType.CODE:
                cell.insert('end', elem.content, ('code',))
            elif elem.type == InlineType.LINEBREAK:
                cell.insert('end', '\n')
            elif elem.type == InlineType.MATH:
                cell.insert('end', elem.content)  # 简化处理
            elif elem.type == InlineType.LINK:
                cell.insert('end', elem.content)
            else:
                cell.insert('end', elem.content if elem.content else '')
    
    def _insert_quote(self, text: str):
        """插入引用"""
        self.text.insert('end', '│ ', ('quote',))
        self._insert_inline_elements(text, extra_tags=['quote'])
        self.text.insert('end', '\n\n')
    
    def _insert_list(self, items: list, ordered: bool = False):
        """插入列表（支持任务列表、有序/无序列表和多级缩进）"""
        # 多级编号计数器: {级别: 当前编号}
        level_counters = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        prev_level = -1
        
        # 多级编号格式
        def get_number_format(level: int, count: int) -> str:
            if level == 0:
                return f'{count}.'
            elif level == 1:
                return f'{chr(ord("a") + (count - 1) % 26)})'
            elif level == 2:
                romans = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x']
                return f'{romans[(count - 1) % 10]}.'
            else:
                return f'{count}.'
        
        for item in items:
            # 获取级别和文本
            if isinstance(item, dict):
                item_level = item.get('level', 0)
                item_text = item.get('text', '')
                item_type = item.get('type', 'item')
                is_task = item_type == 'task'
                checked = item.get('checked', False)
            else:
                item_level = 0
                item_text = str(item)
                is_task = False
                checked = False
            
            # 如果级别变小了，重置更深级别的计数器
            if item_level < prev_level:
                for l in range(item_level + 1, 5):
                    level_counters[l] = 0
            
            # 增加当前级别计数
            level_counters[item_level] += 1
            prev_level = item_level
            
            # 计算缩进
            indent = '  ' + '    ' * item_level
            
            if is_task:
                checkbox = '☑' if checked else '☐'
                self.text.insert('end', f'{indent}{checkbox} ', ('list_item',))
                self._insert_inline_elements(item_text, extra_tags=['list_item'])
            else:
                if ordered:
                    number = get_number_format(item_level, level_counters[item_level])
                    self.text.insert('end', f'{indent}{number} ', ('list_item',))
                else:
                    bullets = ['•', '◦', '▪', '‣']
                    bullet = bullets[min(item_level, len(bullets) - 1)]
                    self.text.insert('end', f'{indent}{bullet} ', ('list_item',))
                self._insert_inline_elements(item_text, extra_tags=['list_item'])
            
            self.text.insert('end', '\n')
        self.text.insert('end', '\n')
    
    def _insert_image(self, alt: str, url: str):
        """插入图片占位"""
        self.text.insert('end', f'🖼️ [{alt}]\n\n')


class SearchReplaceDialog(ctk.CTkToplevel):
    """搜索替换对话框"""
    def __init__(self, master, text_widget):
        super().__init__(master)
        self.text_widget = text_widget
        self.title("🔍 搜索和替换")
        self.geometry("450x200")
        self.resizable(False, False)
        self.transient(master)
        
        # 居中显示
        self.update_idletasks()
        x = master.winfo_x() + (master.winfo_width() - 450) // 2
        y = master.winfo_y() + (master.winfo_height() - 200) // 2
        self.geometry(f"+{x}+{y}")
        
        self.current_match = 0
        self.matches = []
        
        self._create_ui()
        self.search_entry.focus()
    
    def _create_ui(self):
        """创建界面"""
        main_frame = ctk.CTkFrame(self, fg_color=COLORS['bg_card'])
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 搜索行
        search_frame = ctk.CTkFrame(main_frame, fg_color='transparent')
        search_frame.pack(fill='x', pady=(0, 10))
        
        ctk.CTkLabel(search_frame, text="搜索:", width=60).pack(side='left')
        self.search_entry = ctk.CTkEntry(search_frame, width=250)
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind('<Return>', lambda e: self.find_next())
        
        ctk.CTkButton(
            search_frame, text="下一个", width=70,
            command=self.find_next
        ).pack(side='left', padx=2)
        
        # 替换行
        replace_frame = ctk.CTkFrame(main_frame, fg_color='transparent')
        replace_frame.pack(fill='x', pady=(0, 10))
        
        ctk.CTkLabel(replace_frame, text="替换:", width=60).pack(side='left')
        self.replace_entry = ctk.CTkEntry(replace_frame, width=250)
        self.replace_entry.pack(side='left', padx=5)
        
        ctk.CTkButton(
            replace_frame, text="替换", width=70,
            command=self.replace_one
        ).pack(side='left', padx=2)
        
        # 操作按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color='transparent')
        btn_frame.pack(fill='x', pady=(10, 0))
        
        ctk.CTkButton(
            btn_frame, text="全部替换", width=100,
            fg_color=COLORS['warning'],
            command=self.replace_all
        ).pack(side='left', padx=5)
        
        self.status_label = ctk.CTkLabel(
            btn_frame, text="", text_color=COLORS['text_secondary']
        )
        self.status_label.pack(side='left', padx=20)
        
        ctk.CTkButton(
            btn_frame, text="关闭", width=80,
            fg_color=COLORS['text_secondary'],
            command=self.destroy
        ).pack(side='right')
    
    def find_next(self):
        """查找下一个匹配"""
        search_text = self.search_entry.get()
        if not search_text:
            return
        
        # 清除之前的高亮
        self.text_widget.tag_remove('search_highlight', '1.0', 'end')
        
        # 配置高亮标签
        self.text_widget.tag_configure('search_highlight', background=COLORS['highlight'])
        
        # 搜索所有匹配
        self.matches = []
        start_pos = '1.0'
        while True:
            pos = self.text_widget.search(search_text, start_pos, 'end', nocase=True)
            if not pos:
                break
            end_pos = f"{pos}+{len(search_text)}c"
            self.matches.append((pos, end_pos))
            self.text_widget.tag_add('search_highlight', pos, end_pos)
            start_pos = end_pos
        
        if self.matches:
            # 跳转到下一个
            self.current_match = (self.current_match + 1) % len(self.matches)
            pos, end_pos = self.matches[self.current_match]
            self.text_widget.see(pos)
            self.text_widget.mark_set('insert', pos)
            self.status_label.configure(text=f"找到 {len(self.matches)} 个匹配")
        else:
            self.status_label.configure(text="未找到匹配项")
    
    def replace_one(self):
        """替换当前匹配"""
        if not self.matches:
            self.find_next()
            return
        
        search_text = self.search_entry.get()
        replace_text = self.replace_entry.get()
        
        if self.matches:
            pos, end_pos = self.matches[self.current_match]
            self.text_widget.delete(pos, end_pos)
            self.text_widget.insert(pos, replace_text)
            self.find_next()
    
    def replace_all(self):
        """替换所有匹配"""
        search_text = self.search_entry.get()
        replace_text = self.replace_entry.get()
        
        if not search_text:
            return
        
        content = self.text_widget.get('1.0', 'end-1c')
        count = content.count(search_text)
        
        if count > 0:
            new_content = content.replace(search_text, replace_text)
            self.text_widget.delete('1.0', 'end')
            self.text_widget.insert('1.0', new_content)
            self.status_label.configure(text=f"已替换 {count} 处")
            self.matches = []
        else:
            self.status_label.configure(text="未找到匹配项")


class OutlineView(ctk.CTkFrame):
    """大纲视图 - 显示文档标题结构"""
    def __init__(self, master, on_heading_click=None, **kwargs):
        super().__init__(master, fg_color=COLORS['bg_sidebar'], **kwargs)
        
        self.on_heading_click = on_heading_click
        self.headings = []
        
        # 标题
        title_frame = ctk.CTkFrame(self, fg_color='transparent')
        title_frame.pack(fill='x', padx=15, pady=(15, 10))
        
        ctk.CTkLabel(
            title_frame,
            text="📝 大纲",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLORS['text_primary']
        ).pack(side='left')
        
        # 大纲列表
        self.outline_frame = ctk.CTkScrollableFrame(
            self, fg_color='transparent', corner_radius=0
        )
        self.outline_frame.pack(fill='both', expand=True, padx=5)
    
    def update_outline(self, markdown_text: str):
        """更新大纲"""
        # 清除旧内容
        for widget in self.outline_frame.winfo_children():
            widget.destroy()
        
        self.headings = []
        
        # 解析标题
        lines = markdown_text.split('\n')
        for i, line in enumerate(lines):
            match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()
                # 清除Markdown标记
                title = re.sub(r'\*\*(.+?)\*\*', r'\1', title)
                title = re.sub(r'\*(.+?)\*', r'\1', title)
                
                self.headings.append((level, title, i + 1))
                
                # 创建标题按钮
                indent = '  ' * (level - 1)
                btn_text = f"{indent}{'#' * level} {title}"
                if len(btn_text) > 30:
                    btn_text = btn_text[:27] + '...'
                
                btn = ctk.CTkButton(
                    self.outline_frame,
                    text=btn_text,
                    anchor='w',
                    fg_color='transparent',
                    text_color=COLORS['text_primary'] if level <= 2 else COLORS['text_secondary'],
                    hover_color=COLORS['border'],
                    font=ctk.CTkFont(size=12 if level <= 2 else 11),
                    height=28,
                    command=lambda ln=i+1: self._on_click(ln)
                )
                btn.pack(fill='x', pady=1)
    
    def _on_click(self, line_number: int):
        """点击标题时跳转"""
        if self.on_heading_click:
            self.on_heading_click(line_number)


class RecentFilesView(ctk.CTkFrame):
    """最近文件视图"""
    def __init__(self, master, on_file_click=None, **kwargs):
        super().__init__(master, fg_color=COLORS['bg_sidebar'], **kwargs)
        
        self.on_file_click = on_file_click
        
        # 标题
        title_frame = ctk.CTkFrame(self, fg_color='transparent')
        title_frame.pack(fill='x', padx=15, pady=(15, 10))
        
        ctk.CTkLabel(
            title_frame,
            text="📁 最近文件",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLORS['text_primary']
        ).pack(side='left')
        
        # 文件列表
        self.files_frame = ctk.CTkScrollableFrame(
            self, fg_color='transparent', corner_radius=0
        )
        self.files_frame.pack(fill='both', expand=True, padx=5)
    
    def update_files(self, files: list):
        """更新文件列表"""
        # 清除旧内容
        for widget in self.files_frame.winfo_children():
            widget.destroy()
        
        for filepath in files[:10]:  # 最多显示10个
            if os.path.exists(filepath):
                filename = os.path.basename(filepath)
                
                btn = ctk.CTkButton(
                    self.files_frame,
                    text=f"📄 {filename}",
                    anchor='w',
                    fg_color='transparent',
                    text_color=COLORS['text_primary'],
                    hover_color=COLORS['border'],
                    font=ctk.CTkFont(size=12),
                    height=28,
                    command=lambda fp=filepath: self._on_click(fp)
                )
                btn.pack(fill='x', pady=1)
    
    def _on_click(self, filepath: str):
        """点击文件"""
        if self.on_file_click:
            self.on_file_click(filepath)


class App(ctk.CTk):
    """主应用窗口 - 优化版"""
    
    def __init__(self):
        super().__init__()
        
        # 窗口配置
        self.title("✨ Markdown → Word 转换器 v2.0")
        self.geometry("1500x900")
        self.minsize(1100, 700)
        
        # 设置窗口图标
        try:
            import os
            icon_path = os.path.join(os.path.dirname(__file__), 'app.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # 图标加载失败不影响程序运行
        
        # 设置窗口背景
        self.configure(fg_color=COLORS['bg_light'])
        
        # 加载配置
        self.config = load_config()
        
        # 当前文件路径
        self.current_file = None
        self.converter = MarkdownToWordConverter()
        
        # 防抖定时器ID
        self._debounce_id = None
        
        # 搜索对话框引用
        self.search_dialog = None
        
        # 内容修改标记
        self._content_modified = False
        self._last_saved_content = ""
        
        # 自动保存配置
        self._auto_save_interval = 30000  # 30秒
        self._auto_save_id = None
        self._auto_save_file = os.path.join(tempfile.gettempdir(), 'md2word_autosave.md')
        
        # 构建界面
        self._create_header()
        self._create_status_bar()  # 先创建状态栏
        self._create_main_content()  # 再创建主内容（包含_insert_example调用）
        
        # 绑定快捷键
        self.bind('<Control-o>', lambda e: self.open_file())
        self.bind('<Control-s>', lambda e: self.save_file())  # 保存源文件
        self.bind('<Control-Shift-s>', lambda e: self.export_to_word())  # 导出Word
        self.bind('<Control-Shift-c>', lambda e: self.copy_to_clipboard())
        self.bind('<Control-f>', lambda e: self.show_search_dialog())
        self.bind('<Control-h>', lambda e: self.show_search_dialog())
        self.bind('<Control-plus>', lambda e: self.change_font_size(1))
        self.bind('<Control-minus>', lambda e: self.change_font_size(-1))
        self.bind('<Control-b>', lambda e: self.toggle_sidebar())
        self.bind('<Control-p>', lambda e: self.toggle_preview())
        self.bind('<Control-z>', lambda e: self._undo())
        self.bind('<Control-y>', lambda e: self._redo())
        self.bind('<Control-Shift-z>', lambda e: self._redo())
        self.bind('<F1>', lambda e: self.show_help())
        
        # 绑定窗口关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 支持拖拽导入文件
        self._setup_drag_drop()
        
        # 更新最近文件列表
        self._update_recent_files_view()
        
        # 恢复窗口位置和大小
        self._restore_window_geometry()
        
        # 启动自动保存
        self._start_auto_save()
    
    def _create_header(self):
        """创建顶部标题栏"""
        header = ctk.CTkFrame(self, fg_color=COLORS['primary'], height=60, corner_radius=0)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        # 左侧Logo和标题
        left_frame = ctk.CTkFrame(header, fg_color="transparent")
        left_frame.pack(side="left", padx=20, pady=12)
        
        title_label = ctk.CTkLabel(
            left_frame,
            text="📝 Markdown → Word",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        title_label.pack(side="left")
        

        # 中间工具栏
        toolbar_frame = ctk.CTkFrame(header, fg_color="transparent")
        toolbar_frame.pack(side="left", padx=30)
        
        # 工具按钮
        tools = [
            ("📂", "打开", self.open_file, "Ctrl+O"),
            ("💾", "保存", self.export_to_word, "Ctrl+S"),
            ("🔍", "搜索", self.show_search_dialog, "Ctrl+F"),
            ("👁", "预览", self.toggle_preview, "Ctrl+P"),
        ]
        
        for icon, tip, cmd, shortcut in tools:
            btn = ctk.CTkButton(
                toolbar_frame,
                text=icon,
                width=36,
                height=32,
                corner_radius=8,
                fg_color="#8B8CF2",
                hover_color="#9FA0F5",
                text_color="white",
                font=ctk.CTkFont(size=16),
                command=cmd
            )
            btn.pack(side="left", padx=3)
        
        # 插入按钮（带下拉菜单）
        insert_btn = ctk.CTkButton(
            toolbar_frame,
            text="➕",
            width=36,
            height=32,
            corner_radius=8,
            fg_color="#10B981",
            hover_color="#059669",
            text_color="white",
            font=ctk.CTkFont(size=16),
            command=lambda: None  # 占位
        )
        insert_btn.pack(side="left", padx=3)
        insert_btn.bind('<Button-1>', self.show_insert_menu)
        
        # 右侧按钮组
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right", padx=20, pady=12)
        
        # 侧边栏切换
        self.sidebar_btn = ctk.CTkButton(
            btn_frame,
            text="☰",
            command=self.toggle_sidebar,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=8,
            width=36,
            height=32,
            font=ctk.CTkFont(size=16)
        )
        self.sidebar_btn.pack(side="left", padx=3)
        
        # 字体调整
        ctk.CTkButton(
            btn_frame,
            text="A-",
            command=lambda: self.change_font_size(-1),
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=8,
            width=36,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=1)
        
        ctk.CTkButton(
            btn_frame,
            text="A+",
            command=lambda: self.change_font_size(1),
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=8,
            width=36,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=1)
        
        # 主题切换
        self.theme_btn = ctk.CTkButton(
            btn_frame,
            text="🌙",
            command=self.toggle_theme,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=8,
            width=36,
            height=32
        )
        self.theme_btn.pack(side="left", padx=3)
    
    def _create_main_content(self):
        """创建主内容区域 - 包含侧边栏"""
        # 主容器
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 左侧边栏（大纲+最近文件）
        self.sidebar_visible = self.config.get('sidebar_visible', True)
        self.sidebar = ctk.CTkFrame(self.main_container, fg_color=COLORS['bg_sidebar'], width=250, corner_radius=12)
        if self.sidebar_visible:
            self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)
        
        # 侧边栏内容
        self._create_sidebar_content()
        
        # 右侧主编辑区
        self.main_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.main_frame.pack(side="left", fill="both", expand=True)
        
        # 配置列权重：左侧输入略宽，右侧预览略窄
        self.main_frame.grid_columnconfigure(0, weight=3)
        self.main_frame.grid_columnconfigure(1, weight=2)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # ===== 左侧：输入区域 =====
        self._create_input_panel(self.main_frame)
        
        # ===== 右侧：预览区域 =====
        self._create_preview_panel(self.main_frame)
        
        # 插入示例文本（在所有组件创建完成后）
        self._insert_example()
    
    def _create_sidebar_content(self):
        """创建侧边栏内容"""
        # 大纲视图
        self.outline_view = OutlineView(
            self.sidebar,
            on_heading_click=self._jump_to_line
        )
        self.outline_view.pack(fill="both", expand=True, pady=(0, 10))
        
        # 分隔线
        separator = ctk.CTkFrame(self.sidebar, height=1, fg_color=COLORS['border'])
        separator.pack(fill="x", padx=15, pady=5)
        
        # 最近文件
        self.recent_files_view = RecentFilesView(
            self.sidebar,
            on_file_click=self._open_recent_file
        )
        self.recent_files_view.pack(fill="both", expand=True)
    
    def _create_input_panel(self, parent):
        """创建输入面板 - 带行号"""
        self.input_card = ModernCard(parent)
        self.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        
        # 工具栏 - 紧凑布局，紧贴文本框
        toolbar = ctk.CTkFrame(self.input_card, fg_color="transparent", height=26)
        toolbar.pack(fill="x", padx=6, pady=(6, 0))
        toolbar.pack_propagate(False)  # 保持固定高度
        
        # 快捷插入按钮 - 分组显示
        groups = [
            # 标题组
            [("H1", "# "), ("H2", "## "), ("H3", "### ")],
            # 格式组
            [("B", "**粗体**"), ("I", "*斜体*"), ("~", "~~删除~~")],
            # 上下标组
            [("²", "<sup>上标</sup>"), ("₂", "<sub>下标</sub>")],
            # 插入组
            [("🖼", "![图片](url)"), ("🔗", "[链接](url)"), ("∑", "$公式$")],
            # 块级组
            [("≣", "| 表头 |\n|---|\n| 内容 |"), ("`", "```python\ncode\n```")],
        ]
        
        for i, group in enumerate(groups):
            if i > 0:
                # 分隔线
                sep = ctk.CTkFrame(toolbar, width=1, fg_color=COLORS['border'])
                sep.pack(side="left", fill="y", padx=3, pady=2)
            
            for text, insert_text in group:
                btn = ctk.CTkButton(
                    toolbar,
                    text=text,
                    width=26,
                    height=22,
                    corner_radius=4,
                    fg_color=COLORS['bg_light'],
                    text_color=COLORS['text_primary'],
                    hover_color=COLORS['border'],
                    font=ctk.CTkFont(size=10, weight="bold"),
                    command=lambda t=insert_text: self.insert_text(t)
                )
                btn.pack(side="left", padx=1)
        
        # 带行号的输入文本框
        self.input_editor = LineNumberedText(
            self.input_card,
            font_size=self.config.get('font_size', 14),
            on_scroll=self._on_editor_scroll  # 滚动同步回调
        )
        self.input_editor.pack(fill="both", expand=True, padx=6, pady=(4, 6))
        
        # 兼容旧属性名
        self.input_text = self.input_editor
        
        # 绑定实时预览（带防抖）
        self.input_editor.bind('<KeyRelease>', self._on_text_change_debounced)
    
    def _create_preview_panel(self, parent):
        """创建预览面板 - 支持开关"""
        self.preview_visible = True
        self.preview_card = ModernCard(parent, title="👁️ 实时预览")
        self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        
        # 预览组件
        self.preview = MarkdownPreview(self.preview_card, on_content_change=self._on_preview_change)
        self.preview.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        
        # 底部操作按钮
        btn_frame = ctk.CTkFrame(self.preview_card, fg_color="transparent", height=45)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # 导出Word按钮
        self.export_btn = ModernButton(
            btn_frame,
            text="📄 导出",
            command=self.export_to_word,
            style="primary",
            width=80
        )
        self.export_btn.pack(side="left", padx=(0, 6))
        
        # 复制到剪贴板按钮
        self.copy_btn = ModernButton(
            btn_frame,
            text="📋 复制",
            command=self.copy_to_clipboard,
            style="secondary",
            width=80
        )
        self.copy_btn.pack(side="left", padx=(0, 6))
        
        # 关闭预览按钮
        self.hide_preview_btn = ModernButton(
            btn_frame,
            text="✕ 关闭预览",
            command=self.toggle_preview,
            style="outline",
            width=90
        )
        self.hide_preview_btn.pack(side="right")
        
        # 清空按钮
        self.clear_btn = ModernButton(
            btn_frame,
            text="🗑️",
            command=self.clear_all,
            style="outline",
            width=36
        )
        self.clear_btn.pack(side="right", padx=(0, 6))
    
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_bar = ctk.CTkFrame(self, fg_color=COLORS['bg_card'], height=35, corner_radius=0)
        self.status_bar.pack(fill="x", side="bottom")
        
        self.status_label = ctk.CTkLabel(
            self.status_bar,
            text="✨ 就绪 - 支持表格、公式、图片等完整Markdown语法",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary']
        )
        self.status_label.pack(side="left", padx=20, pady=8)
        
        # 字数统计（增强版）
        self.word_count_label = ctk.CTkLabel(
            self.status_bar,
            text="字数: 0 | 行数: 0 | 段落: 0",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary']
        )
        self.word_count_label.pack(side="right", padx=20, pady=8)
    
    def _insert_example(self):
        """插入示例Markdown"""
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
        self.input_text.insert("1.0", example)
        self.on_text_change(None)
    
    def insert_text(self, text: str):
        """在光标位置插入文本"""
        self.input_text.insert("insert", text)
        self.on_text_change(None)
    
    def _on_text_change_debounced(self, event):
        """防抖版文本变化处理 - 300ms延迟"""
        if self._debounce_id:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(300, lambda: self.on_text_change(event))
    
    def on_text_change(self, event):
        """文本变化时更新预览和大纲"""
        content = self.input_text.get("1.0", "end-1c")
        
        # 设置预览区为更新状态（防止循环触发）
        if hasattr(self, 'preview'):
            self.preview.set_updating(True)
            self.preview.update_preview(content)
            self.preview.set_updating(False)
        
        # 更新大纲
        if hasattr(self, 'outline_view'):
            self.outline_view.update_outline(content)
        
        # 更新字数统计（增强版）
        word_count = len(content.replace('\n', '').replace(' ', '').replace('\t', ''))
        line_count = content.count('\n') + 1 if content else 0
        # 统计段落（空行分隔）
        paragraphs = [p for p in content.split('\n\n') if p.strip()]
        para_count = len(paragraphs)
        self.word_count_label.configure(text=f"字数: {word_count} | 行数: {line_count} | 段落: {para_count}")
        
        # 标记内容已修改
        if content != self._last_saved_content:
            self._content_modified = True
            self._update_title()
    
    def _on_preview_change(self, markdown_text: str):
        """预览区内容变化时同步回Markdown编辑器"""
        # 防止循环触发
        if hasattr(self, '_preview_updating') and self._preview_updating:
            return
        
        self._preview_updating = True
        try:
            # 保存当前光标位置
            cursor_pos = self.input_text.text.index(tk.INSERT)
            
            # 更新Markdown编辑器
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", markdown_text)
            
            # 恢复光标位置
            try:
                self.input_text.text.mark_set(tk.INSERT, cursor_pos)
            except:
                pass
            
            # 更新字数统计
            word_count = len(markdown_text.replace('\n', '').replace(' ', '').replace('\t', ''))
            line_count = markdown_text.count('\n') + 1 if markdown_text else 0
            paragraphs = [p for p in markdown_text.split('\n\n') if p.strip()]
            para_count = len(paragraphs)
            self.word_count_label.configure(text=f"字数: {word_count} | 行数: {line_count} | 段落: {para_count}")
            
            # 标记修改
            self._content_modified = True
            self._update_title()
            
            self.update_status("✏️ 预览区已编辑")
        finally:
            self._preview_updating = False
    
    def open_file(self):
        """打开Markdown文件"""
        # 检查未保存的更改
        if not self._check_unsaved_changes():
            return
        
        file_path = filedialog.askopenfilename(
            title="选择Markdown文件",
            filetypes=[
                ("Markdown文件", "*.md *.markdown"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self._load_file(file_path)
    
    def _load_file(self, file_path: str):
        """加载文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", content)
            self.current_file = file_path
            
            # 更新保存状态
            self._last_saved_content = content
            self._content_modified = False
            self._update_title()
            
            self.on_text_change(None)
            
            # 添加到最近文件
            self._add_recent_file(file_path)
            
            self.update_status(f"✅ 已加载: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件:\n{e}")
    
    def export_to_word(self):
        """导出为Word文档"""
        content = self.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "请先输入Markdown内容")
            return
        
        # 直接导出，使用默认设置
        self._do_export(content, "standard", "a4")
    
    def _show_export_options(self, content: str):
        """显示导出选项对话框"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("导出选项")
        dialog.geometry("400x350")
        dialog.transient(self)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 350) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 标题
        ctk.CTkLabel(
            dialog,
            text="📄 导出设置",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(20, 15))
        
        # 样式选择
        style_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        style_frame.pack(fill="x", padx=30, pady=10)
        
        ctk.CTkLabel(
            style_frame,
            text="文档样式：",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w")
        
        style_var = ctk.StringVar(value="standard")
        
        styles = [
            ("standard", "📘 标准样式 - 宋体/Times New Roman"),
            ("academic", "🎓 学术论文 - 严格的学术格式"),
            ("simple", "✨ 简洁样式 - 干净简约")
        ]
        
        for value, label in styles:
            ctk.CTkRadioButton(
                style_frame,
                text=label,
                variable=style_var,
                value=value,
                font=ctk.CTkFont(size=12)
            ).pack(anchor="w", pady=5, padx=10)
        
        # 页面设置
        page_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        page_frame.pack(fill="x", padx=30, pady=10)
        
        ctk.CTkLabel(
            page_frame,
            text="页面设置：",
            font=ctk.CTkFont(size=14)
        ).pack(anchor="w")
        
        page_var = ctk.StringVar(value="a4")
        page_options = ctk.CTkFrame(page_frame, fg_color="transparent")
        page_options.pack(fill="x", pady=5, padx=10)
        
        ctk.CTkRadioButton(page_options, text="A4", variable=page_var, value="a4").pack(side="left", padx=10)
        ctk.CTkRadioButton(page_options, text="Letter", variable=page_var, value="letter").pack(side="left", padx=10)
        
        # 按钮
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=30, pady=20)
        
        def do_export():
            dialog.destroy()
            self._do_export(content, style_var.get(), page_var.get())
        
        ctk.CTkButton(
            btn_frame,
            text="📤 导出",
            command=do_export,
            fg_color=COLORS['primary'],
            width=120
        ).pack(side="right", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="取消",
            command=dialog.destroy,
            fg_color="transparent",
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=80
        ).pack(side="right", padx=5)
    
    def _do_export(self, content: str, style: str, page_size: str):
        """执行导出"""
        # 选择保存路径
        default_name = os.path.splitext(os.path.basename(self.current_file))[0] if self.current_file else "output"
        file_path = filedialog.asksaveasfilename(
            title="保存Word文档",
            defaultextension=".docx",
            initialfile=f"{default_name}.docx",
            filetypes=[("Word文档", "*.docx")]
        )
        
        if file_path:
            self.update_status("⏳ 正在转换...")
            self.export_btn.configure(state="disabled")
            
            # 在线程中执行转换
            def convert():
                try:
                    base_dir = os.path.dirname(self.current_file) if self.current_file else os.getcwd()
                    converter = MarkdownToWordConverter(base_dir=base_dir, style=style, page_size=page_size)
                    converter.convert_text(content)
                    converter.save(file_path)
                    
                    self.after(0, lambda fp=file_path: self.on_export_success(fp))
                except Exception as e:
                    error_msg = str(e)
                    self.after(0, lambda msg=error_msg: self.on_export_error(msg))
            
            threading.Thread(target=convert, daemon=True).start()
    
    def on_export_success(self, file_path):
        """导出成功回调"""
        self.export_btn.configure(state="normal")
        self.update_status(f"✅ 导出成功: {os.path.basename(file_path)}")
        
        if messagebox.askyesno("导出成功", f"文档已保存到:\n{file_path}\n\n是否打开文件？"):
            self._open_file_cross_platform(file_path)
    
    def _open_file_cross_platform(self, file_path: str):
        """跨平台打开文件"""
        import subprocess
        import platform
        
        system = platform.system()
        try:
            if system == 'Windows':
                os.startfile(file_path)
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', file_path], check=True)
            else:  # Linux
                subprocess.run(['xdg-open', file_path], check=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def on_export_error(self, error):
        """导出失败回调"""
        self.export_btn.configure(state="normal")
        self.update_status("❌ 导出失败")
        messagebox.showerror("导出错误", f"转换失败:\n{error}")
    
    def copy_to_clipboard(self):
        """复制内容到剪贴板（Word兼容格式）"""
        content = self.input_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("提示", "请先输入Markdown内容")
            return
        
        self.update_status("⏳ 正在生成剪贴板内容...")
        
        def copy_task():
            try:
                # 生成临时Word文档
                temp_file = tempfile.mktemp(suffix='.docx')
                base_dir = os.path.dirname(self.current_file) if self.current_file else os.getcwd()
                converter = MarkdownToWordConverter(base_dir=base_dir)
                converter.convert_text(content)
                converter.save(temp_file)
                
                # 使用pywin32复制到剪贴板
                self._copy_word_to_clipboard(temp_file)
                
                # 清理临时文件
                os.remove(temp_file)
                
                self.after(0, lambda: self.update_status("✅ 已复制到剪贴板，可直接粘贴到Word"))
                self.after(0, lambda: self._show_copy_toast())
                
            except Exception as e:
                self.after(0, lambda: self.update_status(f"❌ 复制失败: {e}"))
                self.after(0, lambda: messagebox.showerror("错误", f"复制失败:\n{e}"))
        
        threading.Thread(target=copy_task, daemon=True).start()
    
    def _copy_word_to_clipboard(self, docx_path: str):
        """使用COM将Word内容复制到剪贴板"""
        try:
            import win32com.client
            import pythoncom
            
            pythoncom.CoInitialize()
            
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.Content.Copy()  # 复制全部内容（包括公式）
            doc.Close(False)
            word.Quit()
            
            pythoncom.CoUninitialize()
            
        except ImportError:
            # 如果没有pywin32，使用HTML格式复制
            self._copy_as_html(docx_path)
    
    def _copy_as_html(self, docx_path: str):
        """备用方案：转换为HTML复制"""
        content = self.input_text.get("1.0", "end-1c")
        
        # 转换为HTML
        md = md_parser.Markdown(extensions=['tables', 'fenced_code'])
        html = md.convert(content)
        
        # 添加样式
        styled_html = f"""
        <html>
        <head><meta charset="utf-8"></head>
        <body style="font-family: 'Times New Roman', serif; font-size: 11pt;">
        {html}
        </body>
        </html>
        """
        
        self.clipboard_clear()
        self.clipboard_append(styled_html)
    
    def _show_copy_toast(self):
        """显示复制成功提示"""
        toast = ctk.CTkToplevel(self)
        toast.overrideredirect(True)
        toast.attributes('-topmost', True)
        
        # 计算位置
        x = self.winfo_x() + self.winfo_width() // 2 - 150
        y = self.winfo_y() + self.winfo_height() - 100
        toast.geometry(f"300x50+{x}+{y}")
        
        frame = ctk.CTkFrame(toast, fg_color=COLORS['success'], corner_radius=12)
        frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        label = ctk.CTkLabel(
            frame,
            text="✅ 已复制！可直接粘贴到Word",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="white"
        )
        label.pack(expand=True)
        
        # 2秒后自动关闭
        toast.after(2000, toast.destroy)
    
    def clear_all(self):
        """清空所有内容"""
        # 检查未保存的更改
        if not self._check_unsaved_changes():
            return
        
        self.input_text.delete("1.0", "end")
        self.current_file = None
        self._last_saved_content = ""
        self._content_modified = False
        self._update_title()
        self.on_text_change(None)
        self.update_status("✨ 已清空")
    
    def toggle_theme(self):
        """切换明暗主题"""
        global COLORS
        current = ctk.get_appearance_mode()
        new_mode = "dark" if current == "Light" else "light"
        ctk.set_appearance_mode(new_mode)
        
        # 更新颜色配置
        if new_mode == "dark":
            COLORS = COLORS_DARK.copy()
        else:
            COLORS = COLORS_LIGHT.copy()
        
        # 更新按钮图标
        self.theme_btn.configure(text="☀️" if new_mode == "dark" else "🌙")
        
        # 更新窗口背景
        self.configure(fg_color=COLORS['bg_light'])
        
        # 更新编辑器组件颜色
        self._update_editor_theme()
        
        # 更新预览区颜色
        if hasattr(self, 'preview'):
            self._update_preview_theme()
        
        # 更新侧边栏颜色
        if hasattr(self, 'sidebar'):
            self._update_sidebar_theme()
        
        # 更新卡片颜色
        self._update_cards_theme()
    
    def _update_editor_theme(self):
        """更新编辑器主题颜色"""
        try:
            # 更新编辑器容器
            if hasattr(self, 'input_text'):
                self.input_text.container.configure(bg=COLORS['bg_light'])
                self.input_text.text_frame.configure(bg=COLORS['bg_light'])
                # 更新行号区域
                self.input_text.line_numbers.configure(
                    background=COLORS['line_number_bg'],
                    foreground=COLORS['line_number']
                )
                # 更新文本区域
                self.input_text.text.configure(
                    bg=COLORS['editor_bg'],
                    fg=COLORS['text_primary'],
                    insertbackground=COLORS['text_primary']
                )
        except Exception:
            pass
    
    def _update_preview_theme(self):
        """更新预览区主题颜色"""
        try:
            if hasattr(self, 'preview') and self.preview:
                # 更新预览区背景和文字颜色
                self.preview.text.configure(
                    bg=COLORS['preview_bg'],
                    fg=COLORS['text_primary']
                )
                # 更新代码块样式（始终白色底色）
                is_dark = ctk.get_appearance_mode() == "Dark"
                code_bg = '#F5F5F5'  # 浅灰底色
                code_block_bg = '#FAFAFA'  # 浅灰底色
                code_block_fg = '#1F2937'  # 深色文字
                link_color = '#60A5FA' if is_dark else '#0000FF'
                quote_color = '#9CA3AF' if is_dark else '#6B7280'
                math_color = COLORS['text_primary']
                
                self.preview.text.tag_configure('code', background=code_bg)
                self.preview.text.tag_configure('code_block', background=code_block_bg, foreground=code_block_fg)
                self.preview.text.tag_configure('link', foreground=link_color)
                self.preview.text.tag_configure('quote', foreground=quote_color)
                self.preview.text.tag_configure('math', foreground=math_color)
                self.preview.text.tag_configure('math_block', foreground=math_color)
        except Exception:
            pass
    
    def _update_sidebar_theme(self):
        """更新侧边栏主题颜色"""
        try:
            if hasattr(self, 'sidebar') and self.sidebar:
                self.sidebar.configure(fg_color=COLORS['bg_sidebar'])
        except Exception:
            pass
    
    def _update_cards_theme(self):
        """更新卡片组件主题颜色"""
        try:
            # 更新输入卡片
            if hasattr(self, 'input_card'):
                self.input_card.configure(fg_color=COLORS['bg_card'], border_color=COLORS['border'])
            # 更新预览卡片
            if hasattr(self, 'preview_card'):
                self.preview_card.configure(fg_color=COLORS['bg_card'], border_color=COLORS['border'])
        except Exception:
            pass
    
    def update_status(self, message: str):
        """更新状态栏"""
        self.status_label.configure(text=message)
    
    # ==================== 新增功能方法 ====================
    
    def toggle_preview(self):
        """切换预览显示/隐藏"""
        self.preview_visible = not self.preview_visible
        
        if self.preview_visible:
            # 显示预览
            self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
            self.hide_preview_btn.configure(text="✕ 关闭预览")
            # 调整列权重
            self.main_frame.grid_columnconfigure(0, weight=3)
            self.main_frame.grid_columnconfigure(1, weight=2)
            # 更新预览
            self.on_text_change(None)
            self.update_status("👁️ 预览已开启")
        else:
            # 隐藏预览
            self.preview_card.grid_forget()
            # 调整输入区域占满
            self.main_frame.grid_columnconfigure(0, weight=1)
            self.main_frame.grid_columnconfigure(1, weight=0)
            self.update_status("📝 纯编辑模式 - 按 Ctrl+P 或点击工具栏打开预览")
    
    def toggle_sidebar(self):
        """切换侧边栏显示/隐藏"""
        self.sidebar_visible = not self.sidebar_visible
        
        if self.sidebar_visible:
            self.sidebar.pack(side="left", fill="y", padx=(0, 10), before=self.main_container.winfo_children()[1])
        else:
            self.sidebar.pack_forget()
        
        # 保存配置
        self.config['sidebar_visible'] = self.sidebar_visible
        save_config(self.config)
    
    def change_font_size(self, delta: int):
        """调整字体大小"""
        current_size = self.config.get('font_size', 14)
        new_size = max(10, min(24, current_size + delta))
        
        if new_size != current_size:
            self.config['font_size'] = new_size
            save_config(self.config)
            
            # 更新编辑器字体
            if hasattr(self, 'input_editor'):
                self.input_editor.set_font_size(new_size)
            
            self.update_status(f"🔤 字体大小: {new_size}px")
    
    def show_search_dialog(self):
        """显示搜索替换对话框"""
        if self.search_dialog is None or not self.search_dialog.winfo_exists():
            # 获取实际的text widget
            text_widget = self.input_editor._textbox
            self.search_dialog = SearchReplaceDialog(self, text_widget)
        else:
            self.search_dialog.focus()
    
    def _setup_drag_drop(self):
        """设置拖拽导入支持"""
        try:
            # 尝试使用tkinterdnd2（如果安装了）
            from tkinterdnd2 import DND_FILES, TkinterDnD
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop)
        except ImportError:
            # 没有安装tkinterdnd2，使用简单的方式
            pass
    
    def _on_drop(self, event):
        """处理拖拽放置事件"""
        file_path = event.data
        # 清理路径（去除大括号等）
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        if file_path.lower().endswith(('.md', '.markdown', '.txt')):
            self._load_file(file_path)
        else:
            messagebox.showwarning("提示", "请拖拽Markdown文件(.md, .markdown, .txt)")
    
    def _add_recent_file(self, file_path: str):
        """添加到最近文件列表"""
        recent = self.config.get('recent_files', [])
        
        # 如果已存在，先移除
        if file_path in recent:
            recent.remove(file_path)
        
        # 添加到开头
        recent.insert(0, file_path)
        
        # 最多保留10个
        self.config['recent_files'] = recent[:10]
        save_config(self.config)
        
        # 更新视图
        self._update_recent_files_view()
    
    def _update_recent_files_view(self):
        """更新最近文件视图"""
        if hasattr(self, 'recent_files_view'):
            recent = self.config.get('recent_files', [])
            self.recent_files_view.update_files(recent)
    
    def _open_recent_file(self, file_path: str):
        """打开最近文件"""
        if os.path.exists(file_path):
            self._load_file(file_path)
        else:
            messagebox.showwarning("提示", f"文件不存在:\n{file_path}")
            # 从列表中移除
            recent = self.config.get('recent_files', [])
            if file_path in recent:
                recent.remove(file_path)
                self.config['recent_files'] = recent
                save_config(self.config)
                self._update_recent_files_view()
    
    def _jump_to_line(self, line_number: int):
        """跳转到指定行"""
        try:
            # 设置光标位置
            index = f"{line_number}.0"
            self.input_text._textbox.see(index)
            self.input_text._textbox.mark_set("insert", index)
            self.input_text._textbox.focus()
        except Exception:
            pass
    
    # ==================== 文件保存功能 ====================
    
    def save_file(self):
        """保存Markdown源文件"""
        if self.current_file:
            self._save_to_file(self.current_file)
        else:
            self.save_file_as()
    
    def save_file_as(self):
        """另存为Markdown文件"""
        file_path = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            initialfile="untitled.md",
            filetypes=[
                ("Markdown文件", "*.md"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self._save_to_file(file_path)
            self.current_file = file_path
            self._add_recent_file(file_path)
    
    def _save_to_file(self, file_path: str):
        """实际保存文件"""
        try:
            content = self.input_text.get("1.0", "end-1c")
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            self._last_saved_content = content
            self._content_modified = False
            self._update_title()
            self.update_status(f"✅ 已保存: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存文件:\n{e}")
    
    def _check_unsaved_changes(self) -> bool:
        """检查未保存的更改，返回 True 表示可以继续操作"""
        if not self._content_modified:
            return True
        
        current_content = self.input_text.get("1.0", "end-1c")
        if current_content == self._last_saved_content:
            return True
        
        result = messagebox.askyesnocancel(
            "未保存的更改",
            "当前文档有未保存的更改。\n\n是否保存？"
        )
        
        if result is None:  # 取消
            return False
        elif result:  # 是 - 保存
            self.save_file()
            return True
        else:  # 否 - 不保存
            return True
    
    def _on_closing(self):
        """窗口关闭事件"""
        if self._check_unsaved_changes():
            # 保存窗口位置和大小
            self._save_window_geometry()
            self.destroy()
    
    def _update_title(self):
        """更新窗口标题"""
        base_title = "✨ Markdown → Word 转换器 by 一个好人"
        if self.current_file:
            filename = os.path.basename(self.current_file)
            modified = " *" if self._content_modified else ""
            self.title(f"{filename}{modified} - {base_title}")
        else:
            modified = " *" if self._content_modified else ""
            self.title(f"未命名{modified} - {base_title}")
    
    # ==================== 撤销重做 ====================
    
    def _undo(self):
        """撤销操作"""
        try:
            self.input_text._textbox.edit_undo()
            self.on_text_change(None)
        except tk.TclError:
            pass  # 没有可撤销的操作
    
    def _redo(self):
        """重做操作"""
        try:
            self.input_text._textbox.edit_redo()
            self.on_text_change(None)
        except tk.TclError:
            pass  # 没有可重做的操作
    
    # ==================== 窗口位置记忆 ====================
    
    def _save_window_geometry(self):
        """保存窗口位置和大小"""
        self.config['window_geometry'] = self.geometry()
        save_config(self.config)
    
    def _restore_window_geometry(self):
        """恢复窗口位置和大小"""
        geometry = self.config.get('window_geometry')
        if geometry:
            try:
                self.geometry(geometry)
            except Exception:
                pass  # 如果恢复失败，使用默认尺寸
    
    # ==================== 帮助菜单 ====================
    
    def show_help(self):
        """显示快捷键帮助"""
        help_dialog = ctk.CTkToplevel(self)
        help_dialog.title("⌨️ 快捷键说明")
        help_dialog.geometry("400x450")
        help_dialog.transient(self)
        help_dialog.resizable(False, False)
        
        # 居中显示
        help_dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 450) // 2
        help_dialog.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(help_dialog, fg_color=COLORS['bg_card'])
        frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(
            frame,
            text="快捷键说明",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COLORS['text_primary']
        ).pack(pady=(10, 20))
        
        shortcuts = [
            ("Ctrl+O", "打开文件"),
            ("Ctrl+S", "保存Markdown文件"),
            ("Ctrl+Shift+S", "导出为Word文档"),
            ("Ctrl+Shift+C", "复制到剪贴板"),
            ("Ctrl+Z", "撤销"),
            ("Ctrl+Y", "重做"),
            ("Ctrl+F", "搜索替换"),
            ("Ctrl+P", "切换预览"),
            ("Ctrl+B", "切换侧边栏"),
            ("Ctrl++/-", "调整字体大小"),
            ("F1", "显示此帮助"),
        ]
        
        for key, desc in shortcuts:
            row = ctk.CTkFrame(frame, fg_color='transparent')
            row.pack(fill='x', pady=3, padx=20)
            
            ctk.CTkLabel(
                row,
                text=key,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=COLORS['primary'],
                width=120,
                anchor='w'
            ).pack(side='left')
            
            ctk.CTkLabel(
                row,
                text=desc,
                font=ctk.CTkFont(size=12),
                text_color=COLORS['text_secondary'],
                anchor='w'
            ).pack(side='left', fill='x', expand=True)
        
        ctk.CTkButton(
            frame,
            text="确定",
            command=help_dialog.destroy,
            fg_color=COLORS['primary'],
            width=100
        ).pack(pady=20)
    
    # ==================== 同步滚动 ====================
    
    def _on_editor_scroll(self, position: float):
        """编辑器滚动时同步预览区"""
        if hasattr(self, 'preview') and self.preview_visible:
            self.preview.text.yview_moveto(position)
    
    # ==================== 自动保存 ====================
    
    def _start_auto_save(self):
        """启动自动保存定时器"""
        self._do_auto_save()
    
    def _do_auto_save(self):
        """执行自动保存"""
        try:
            content = self.input_text.get("1.0", "end-1c")
            if content.strip():  # 只有内容不为空时才保存
                with open(self._auto_save_file, 'w', encoding='utf-8') as f:
                    f.write(content)
        except Exception:
            pass
        finally:
            # 继续下一次自动保存
            self._auto_save_id = self.after(self._auto_save_interval, self._do_auto_save)
    
    def _check_auto_save_recovery(self):
        """检查是否有自动保存的文件可恢复"""
        try:
            if os.path.exists(self._auto_save_file):
                # 检查文件修改时间，如果在10分钟内则提示恢复
                mtime = os.path.getmtime(self._auto_save_file)
                import time
                if time.time() - mtime < 600:  # 10分钟内
                    with open(self._auto_save_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    if content.strip():
                        result = messagebox.askyesno(
                            "恢复自动保存",
                            "发现上次未保存的内容，是否恢复？"
                        )
                        if result:
                            self.input_text.delete("1.0", "end")
                            self.input_text.insert("1.0", content)
                            self.on_text_change(None)
                            self.update_status("✅ 已恢复自动保存的内容")
                            return
                # 删除旧的自动保存文件
                os.remove(self._auto_save_file)
        except Exception:
            pass
    
    def _clear_auto_save(self):
        """清除自动保存文件"""
        try:
            if os.path.exists(self._auto_save_file):
                os.remove(self._auto_save_file)
        except Exception:
            pass
    
    # ==================== 插入菜单 ====================
    
    def show_insert_menu(self, event=None):
        """显示插入菜单 - 使用大对话框"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("插入内容")
        dialog.geometry("500x480")  # 增大尺寸
        dialog.transient(self)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 500) // 2
        y = self.winfo_y() + (self.winfo_height() - 480) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 标题
        ctk.CTkLabel(
            dialog,
            text="➕ 插入内容",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=(20, 15))
        
        # 插入选项按钮容器
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="both", expand=True, padx=30, pady=10)
        
        # 插入选项列表
        insert_options = [
            ("📊 表格", "插入三线表样式表格", self._insert_table_template),
            ("🔗 链接", "插入超链接", self._insert_link_template),
            ("🖼️ 图片", "插入图片引用", self._insert_image_template),
            ("π 公式", "插入LaTeX数学公式", self._insert_math_template),
            ("📝 代码块", "插入代码块", self._insert_code_template),
            ("☐ 任务列表", "插入任务清单", self._insert_task_template),
            ("─── 分割线", "插入水平分割线", lambda: self.insert_text("\n---\n")),
        ]
        
        for i, (icon_text, desc, cmd) in enumerate(insert_options):
            row = ctk.CTkFrame(btn_frame, fg_color="transparent")
            row.pack(fill="x", pady=6)
            
            def make_callback(command):
                def callback():
                    dialog.destroy()
                    command()
                return callback
            
            btn = ctk.CTkButton(
                row,
                text=icon_text,
                font=ctk.CTkFont(size=16),
                width=160,  # 增宽按钮
                height=42,
                fg_color=COLORS['primary'],
                hover_color=COLORS['primary_hover'],
                command=make_callback(cmd)
            )
            btn.pack(side="left", padx=(0, 15))
            
            ctk.CTkLabel(
                row,
                text=desc,
                font=ctk.CTkFont(size=14),
                text_color=COLORS['text_secondary']
            ).pack(side="left", fill="x")
        
        # 关闭按钮
        ctk.CTkButton(
            dialog,
            text="关闭",
            command=dialog.destroy,
            fg_color="transparent",
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=100
        ).pack(pady=20)
    
    def _insert_table_template(self):
        """插入表格模板"""
        template = """| 列一 | 列二 | 列三 |
|------|------|------|
| 内容1 | 内容2 | 内容3 |
| 内容4 | 内容5 | 内容6 |
"""
        self.insert_text(template)
        self.update_status("✅ 已插入表格模板")
    
    def _insert_link_template(self):
        """插入链接模板"""
        self.insert_text("[链接文本](https://example.com)")
        self.update_status("✅ 已插入链接模板")
    
    def _insert_image_template(self):
        """插入图片模板"""
        self.insert_text("![图片描述](图片路径)")
        self.update_status("✅ 已插入图片模板")
    
    def _insert_math_template(self):
        """插入公式模板"""
        template = """$$
\\frac{a}{b} = c
$$"""
        self.insert_text(template)
        self.update_status("✅ 已插入公式模板")
    
    def _insert_code_template(self):
        """插入代码块模板"""
        template = """```python
# 在此输入代码
print("Hello, World!")
```"""
        self.insert_text(template)
        self.update_status("✅ 已插入代码块模板")
    
    def _insert_task_template(self):
        """插入任务列表模板"""
        template = """- [ ] 待完成任务 1
- [ ] 待完成任务 2
- [x] 已完成任务
"""
        self.insert_text(template)
        self.update_status("✅ 已插入任务列表模板")


def main():
    """启动应用"""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
