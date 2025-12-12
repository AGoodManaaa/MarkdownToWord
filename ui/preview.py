# -*- coding: utf-8 -*-

from io import BytesIO
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
import matplotlib.pyplot as plt

from parser import parse_markdown, parse_inline, parse_table, InlineType
from utils import normalize_markdown, convert_latex_delimiters
from ui.theme import COLORS


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
            except Exception:
                pass
        self.text.after(10, center_table)
    
    def _insert_cell_content(self, cell: tk.Text, text: str, is_header: bool = False):
        """在表格单元格中插入格式化内容"""
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
