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
    """可编辑的Markdown预览组件 - 支持双向同步和滚动同步"""
    def __init__(self, master, on_content_change=None, app=None, on_scroll=None, **kwargs):
        super().__init__(master, fg_color=COLORS['bg_card'], corner_radius=12, **kwargs)

        # 预览缩放（仅影响预览区字体/样式）
        self._scale = 1.0
        # 放大基础字号，改善可读性
        self._base_sizes = {
            'body': 18,
            'h1': 30,
            'h2': 24,
            'h3': 20,
            'h4': 18,
            'code': 11,
            'math': 18,
            'math_block': 20,
            'supsub': 10,
            'quote': 12,
            'list_item': 18,
        }
        
        # 内容变化回调
        self.on_content_change = on_content_change
        
        # 滚动同步回调
        self.on_scroll = on_scroll
        self._scroll_updating = False
        # 是否启用与编辑器的滚动同步（默认启用）
        self._sync_scroll_enabled = True

        # App 引用：用于“复制到 Word”
        self.app = app
        
        # 是否正在更新（防止循环触发）
        self._updating = False
        
        # 存储段落信息：{line_start: {'type': 'paragraph', 'md_line': 1, 'format': []}}
        self.paragraph_map = {}

        # 公式 token：用于把“图片公式”安全回写为 Markdown（不改渲染逻辑）
        self._math_token_counter = 0
        self._math_token_map = {}
        
        # 使用 Text widget 支持富文本，整体模拟排版后的页面效果
        # 预览字体改为多字体回退，避免缺字/符号字体
        self.text = tk.Text(
            self,
            wrap='word',
            bg='#FFFFFF',
            fg='#111827',
            font=('Microsoft YaHei', 18, 'normal'),
            width=60,
            padx=20,
            pady=30,
            relief='flat',
            cursor='arrow',  # 预览区只读，但可选中复制
            spacing1=0,
            spacing3=0,
            undo=True,  # 启用撤销
        )
        
        # 滚动条
        self.scrollbar = ctk.CTkScrollbar(self, command=self._on_scrollbar)
        self.text.configure(yscrollcommand=self._on_text_scroll)
        
        self.scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        self.text.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        
        # 配置文本标签样式 - 与Word导出尽量保持一致
        self._setup_tags()
        
        # 存储公式图片引用（防止被垃圾回收）
        self.math_images = []
        
        # 公式计数器
        self.equation_counter = 0
        
        # 预览区只读：允许选中/复制，但禁止输入/粘贴等修改
        self._readonly = True
        self._page_view_enabled = False  # 仿真页面模式开关
        try:
            self.text.configure(state='disabled')
        except Exception:
            pass

        self.text.bind('<Key>', self._block_edit_event)
        self.text.bind('<<Paste>>', self._block_edit_event)
        self.text.bind('<<Cut>>', self._block_edit_event)
        self.text.bind('<Control-v>', self._block_edit_event)
        self.text.bind('<Control-x>', self._block_edit_event)
        self.text.bind('<Control-c>', self._copy_selection)
        self.text.bind('<Control-C>', self._copy_selection)
        self.text.bind('<Control-b>', self._block_edit_event)
        self.text.bind('<Control-i>', self._block_edit_event)
        self.text.bind('<BackSpace>', self._block_edit_event)
        self.text.bind('<Delete>', self._block_edit_event)
        self.text.bind('<Return>', self._block_edit_event)
        self.text.bind('<Control-a>', self._select_all)
        
        # 鼠标滚轮事件绑定（用于滚动同步）
        self.text.bind('<MouseWheel>', self._on_mousewheel)
        self.text.bind('<Button-4>', self._on_mousewheel)  # Linux
        self.text.bind('<Button-5>', self._on_mousewheel)  # Linux
        
        # 双击跳转到源码行
        self.text.bind('<Double-Button-1>', self._on_double_click)
        
        # 跳转回调
        self.on_jump_to_line = None
        
        # 右键菜单
        self._create_context_menu()
    
    def _create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="复制", command=lambda: self.text.event_generate('<<Copy>>'))
        self.context_menu.add_command(label="复制选中到Word（保持格式）", command=self._copy_selection_to_word)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="跳转到编辑器", command=self._jump_to_editor_from_context)
        
        self.text.bind('<Button-3>', self._show_context_menu)

    def _block_edit_event(self, event=None):
        """只读预览区：拦截会修改内容的事件（但不影响选中/复制）。"""
        try:
            if bool(getattr(self, '_readonly', False)):
                # 允许 Ctrl+C 复制（否则会被 <Key> 的拦截吞掉）
                try:
                    if event is not None and (event.state & 0x0004) and str(getattr(event, 'keysym', '')).lower() == 'c':
                        return None
                except Exception:
                    pass
                return 'break'
        except Exception:
            return 'break'
        return None

    def _copy_selection(self, event=None):
        """预览区 Ctrl+C：复制选中内容到剪贴板。"""
        try:
            self.text.event_generate('<<Copy>>')
        except Exception:
            pass
        return 'break'

    def _select_all(self, event=None):
        """支持 Ctrl+A 全选（只读模式下也可用）。"""
        try:
            self.text.tag_add(tk.SEL, '1.0', 'end-1c')
            self.text.mark_set(tk.INSERT, '1.0')
            self.text.see(tk.INSERT)
        except Exception:
            pass
        return 'break'

    def _copy_selection_to_word(self):
        try:
            if self.app is None:
                return

            md = self.get_selection_as_markdown()
            if not (md or '').strip():
                return
            try:
                self.app.copy_markdown_to_clipboard(md)
            except Exception:
                pass
        except Exception:
            pass

    def get_selection_as_markdown(self) -> str:
        """把预览区选中内容转换为 Markdown（含公式 token 还原）。"""
        try:
            sel_first = self.text.index(tk.SEL_FIRST)
            sel_last = self.text.index(tk.SEL_LAST)
        except tk.TclError:
            return ''

        # 扩大一丢丢范围，尽量把紧贴图片公式的隐藏 token 包进来
        start = sel_first
        end = sel_last
        try:
            start = self.text.index(f"{sel_first} -1c")
        except Exception:
            start = sel_first
        try:
            end = self.text.index(f"{sel_last} +1c")
        except Exception:
            end = sel_last

        # 逐行转换，保留行内格式（粗体/斜体/代码/上下标）
        try:
            start_line = int(str(start).split('.')[0])
            start_col = int(str(start).split('.')[1])
            end_line = int(str(end).split('.')[0])
            end_col = int(str(end).split('.')[1])
        except Exception:
            return ''

        lines_out = []
        for ln in range(start_line, end_line + 1):
            if ln == start_line:
                sc = start_col
            else:
                sc = 0

            if ln == end_line:
                ec = end_col
            else:
                try:
                    ec = int(str(self.text.index(f"{ln}.end")).split('.')[1])
                except Exception:
                    ec = 0

            try:
                line_md = self._format_range_line(ln, sc, ec)
            except Exception:
                line_md = ''
            lines_out.append(line_md)

        out = "\n".join(lines_out).strip('\n')
        try:
            out = self._restore_math_tokens(out)
        except Exception:
            pass
        return out

    def _format_range_line(self, line_num: int, start_col: int, end_col: int) -> str:
        if end_col <= start_col:
            return ''

        try:
            raw = self.text.get(f"{line_num}.{start_col}", f"{line_num}.{end_col}")
        except Exception:
            raw = ''
        if raw == '':
            return ''

        # 标题：只在从行首开始选中时才补 #
        prefix = ''
        try:
            if start_col == 0:
                tags0 = set(self.text.tag_names(f"{line_num}.0"))
                if 'h1' in tags0:
                    prefix = '# '
                elif 'h2' in tags0:
                    prefix = '## '
                elif 'h3' in tags0:
                    prefix = '### '
                elif 'h4' in tags0:
                    prefix = '#### '
        except Exception:
            prefix = ''

        segments = []
        current_text = ''
        current_tags = set()

        for offset, ch in enumerate(raw):
            pos = f"{line_num}.{start_col + offset}"
            char_tags = set(self.text.tag_names(pos))
            format_tags = char_tags & {'bold', 'italic', 'strikethrough', 'code', 'superscript', 'subscript', 'math_token'}

            if format_tags != current_tags:
                if current_text:
                    segments.append((current_text, current_tags))
                current_text = ch
                current_tags = format_tags
            else:
                current_text += ch

        if current_text:
            segments.append((current_text, current_tags))

        result = ''
        for text, tags in segments:
            formatted = text
            # token 直接透传，后面统一 restore
            if 'math_token' in tags:
                result += formatted
                continue

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

        return prefix + result
    
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
        # 只有真实编辑（内容被修改）才允许回写，避免点击公式图片/选中导致覆盖 Markdown
        try:
            if not bool(self.text.edit_modified()):
                return
        except Exception:
            pass
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
            try:
                self.text.edit_modified(False)
            except Exception:
                pass
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
        
        md = '\n'.join(result)
        try:
            md = self._restore_math_tokens(md)
        except Exception:
            pass
        return md

    def _new_math_token(self, formula: str, inline: bool) -> str:
        """生成用于回写的“隐藏公式文本”。

        这里不再使用 [[MATH:n]] 占位符，避免占位符泄漏到 Markdown 源文中。
        """
        f = (formula or '').strip()
        if not f:
            return ''
        if inline:
            return f"${f}$"
        return f"$$\n{f}\n$$"

    def _restore_math_tokens(self, text: str) -> str:
        """把 token 还原为 Markdown 公式。"""
        if not text:
            return text
        mp = getattr(self, '_math_token_map', None) or {}
        if not mp:
            # 兼容历史遗留的 [[MATH:n]] 占位符：无法还原时至少清除，避免污染 Markdown 源文
            try:
                import re
                return re.sub(r"\[\[MATH:\d+\]\]", "", text)
            except Exception:
                return text
        out = text
        # token 数量通常不多，直接 replace
        for token, repl in mp.items():
            try:
                out = out.replace(token, repl)
            except Exception:
                pass
        # 再清一次残留的占位符（避免映射不完整时泄漏）
        try:
            import re
            out = re.sub(r"\[\[MATH:\d+\]\]", "", out)
        except Exception:
            pass
        return out
    
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
    
    def set_page_view(self, enabled: bool):
        """切换仿真页面模式"""
        self._page_view_enabled = enabled
        if enabled:
            # 仿真页面样式：固定宽度，居中，带阴影效果（通过边框模拟）
            self.text.pack_forget()
            self.scrollbar.pack_forget()
            
            # 容器背景色（模拟桌面）
            self.configure(fg_color=COLORS.get('bg_sidebar', '#f3f4f6'))
            
            # 重新布局：使用 place 或 grid 居中
            self.text.pack(side="top", pady=20, padx=40, expand=False)
            self.text.configure(width=85, borderwidth=1, relief="solid") # 约 800px
            self.scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
        else:
            # 流式布局样式
            self.text.pack_forget()
            self.scrollbar.pack_forget()
            
            self.configure(fg_color=COLORS['bg_card'])
            self.text.configure(width=60, borderwidth=0, relief="flat")
            
            self.scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
            self.text.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        
        self._setup_tags()

    def _setup_tags(self):
        """配置文本标签样式 - 模拟Word中的样式，更接近真实排版"""
        def sz(key: str) -> int:
            try:
                base = int(self._base_sizes.get(key, 16))
                v = int(round(base * float(self._scale or 1.0)))
                return max(8, min(60, v))
            except Exception:
                return 16

        # 边距也随缩放微调，并按当前预览宽度做封顶，避免挤压导致异常换行
        def margin(base: int, max_frac: float = 0.20) -> int:
            try:
                # 缩放 < 1 时边距减小更多，避免窄屏换行
                s = float(self._scale or 1.0)
                if s < 1.0:
                    factor = 0.5 + 0.5 * s  # 0.9 -> 0.95, 更小时更激进
                else:
                    factor = s
                v = max(4, int(round(base * factor)))

                w = 0
                try:
                    w = int(self.text.winfo_width() or 0)
                except Exception:
                    w = 0
                if w > 0:
                    cap = int(w * float(max_frac))
                    if cap > 0:
                        v = min(v, cap)
                return v
            except Exception:
                return base

        # 更新 Text 默认字体（不依赖 tag 的部分）
        try:
            self.text.configure(font=('宋体', sz('body')))
        except Exception:
            pass

        # 标题样式 - 与 Word 导出保持一致
        # 一级标题：黑体，22pt，居中，段前24pt，段后18pt
        self.text.tag_configure('h1', 
            font=('黑体', sz('h1'), 'bold'), 
            justify='center', 
            spacing1=24,  # 段前
            spacing3=18   # 段后
        )
        # 二级标题：黑体，16pt，居中，段前18pt，段后12pt
        self.text.tag_configure('h2', 
            font=('黑体', sz('h2'), 'bold'), 
            justify='center', 
            spacing1=18, 
            spacing3=12
        )
        # 三级标题：黑体，15pt，左对齐，段前13pt，段后10pt
        self.text.tag_configure('h3', 
            font=('黑体', sz('h3'), 'bold'), 
            lmargin1=0,  # 与 Word 一致，无额外左边距
            lmargin2=0,
            spacing1=13, 
            spacing3=10
        )
        # 四级标题：黑体，14pt，左对齐，段前10pt，段后6pt
        self.text.tag_configure('h4', 
            font=('黑体', sz('h4'), 'bold'), 
            lmargin1=0,  # 与 Word 一致，无额外左边距
            lmargin2=0,
            spacing1=10, 
            spacing3=6
        )
        
        # 正文样式：与 Word 导出保持一致
        # 宋体，12pt（预览中用16px），首行缩进2字符，1.5倍行距，段后6pt
        body_font_size = sz('body')
        first_indent = int(body_font_size * 2)  # 首行缩进2字符（约32px）
        self.text.tag_configure(
            'body',
            font=('宋体', body_font_size),
            lmargin1=first_indent,  # 首行缩进
            lmargin2=0,             # 后续行无额外边距
            rmargin=0,              # 右边距由 Text widget 的 padx 控制
            spacing1=0,    # 段前
            spacing3=6,    # 段后
            spacing2=int(body_font_size * 0.5),  # 行间距（模拟1.5倍行距）
        )
        
        # 粗体、斜体（保持与正文字号一致）
        self.text.tag_configure('bold', font=('宋体', sz('body'), 'bold'))
        self.text.tag_configure('italic', font=('宋体', sz('body'), 'italic'))
        self.text.tag_configure('bold_italic', font=('宋体', sz('body'), 'bold italic'))
        
        # 代码（Consolas，10pt，浅灰背景）
        self.text.tag_configure('code', 
            font=('Consolas', sz('code')), 
            background='#F5F5F5',
            foreground='#1F2937'
        )
        code_indent = int(20 * (self._scale or 1.0))
        self.text.tag_configure('code_block', 
            font=('Consolas', sz('code')), 
            background='#FAFAFA', 
            foreground='#1F2937',
            lmargin1=code_indent,
            lmargin2=code_indent,
            spacing1=6,
            spacing3=6
        )
        
        # 公式：优先使用 Cambria Math，回退到其他数学字体
        math_font = self._get_math_font()
        self.text.tag_configure('math', font=(math_font, sz('math')), foreground='#1a1a2e')
        self.text.tag_configure('math_block', 
            font=(math_font, sz('math_block')), 
            foreground='#1a1a2e', 
            justify='center', 
            spacing1=8, 
            spacing3=8
        )
        
        # 链接（蓝色，下划线）
        self.text.tag_configure('link', foreground='#0000FF', underline=True)
        
        # 删除线
        self.text.tag_configure('strikethrough', overstrike=True)
        
        # 上标和下标
        self.text.tag_configure('superscript', font=('宋体', sz('supsub')), offset=6)
        self.text.tag_configure('subscript', font=('宋体', sz('supsub')), offset=-3)
        
        # 引用（Times New Roman，斜体，左右缩进）
        quote_indent = int(36 * (self._scale or 1.0))
        self.text.tag_configure('quote', 
            font=('Times New Roman', sz('quote'), 'italic'), 
            foreground='#6B7280', 
            lmargin1=quote_indent, 
            lmargin2=quote_indent,
            rmargin=quote_indent,
            spacing1=6,
            spacing3=6
        )
        
        # 列表：与 Word 保持一致的缩进（约1.27cm ≈ 36px）
        list_indent = int(36 * (self._scale or 1.0))
        self.text.tag_configure(
            'list_item',
            font=('宋体', sz('list_item')),
            lmargin1=list_indent,
            lmargin2=list_indent + int(16 * (self._scale or 1.0)),
            spacing1=2,
            spacing3=2,
        )

        # 表格整体居中显示
        self.text.tag_configure('table_block', justify='center')

        # 隐藏公式 token（token 仍会被 Text.get() 取到，用于回写 Markdown）
        try:
            self.text.tag_configure('math_token', elide=True)
        except Exception:
            try:
                self.text.tag_configure('math_token', foreground=self.text.cget('bg'))
            except Exception:
                pass
        
        # 提高上下标标签的优先级，确保字体大小生效
        self.text.tag_raise('superscript')
        self.text.tag_raise('subscript')

    def set_scale(self, scale: float):
        """设置预览缩放比例（仅预览区）。
        
        Args:
            scale: 缩放比例 (0.5 - 1.5)
        """
        try:
            scale = float(scale)
        except Exception:
            scale = 1.0
        # 扩大缩放范围：50% - 150%
        scale = min(1.5, max(0.5, scale))
        if abs(scale - float(getattr(self, '_scale', 1.0))) < 0.01:
            return
        self._scale = scale
        try:
            self._setup_tags()
        except Exception:
            pass
    
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
        """更新预览内容 - 异步版"""
        import threading
        
        # 如果当前已经在渲染，则不再触发新的（由防抖控制频率）
        if getattr(self, '_is_rendering', False):
            return
            
        def _render_task():
            self._is_rendering = True
            try:
                # 预处理和解析在子线程完成
                processed_text = convert_latex_delimiters(markdown_text)
                processed_text = normalize_markdown(processed_text)
                blocks = parse_markdown(processed_text)
                
                # 回到主线程更新 UI
                self.after(0, lambda: self._apply_render_result(blocks))
            except Exception as e:
                print(f"Render error: {e}")
                self._is_rendering = False
        
        threading.Thread(target=_render_task, daemon=True).start()

    def _apply_render_result(self, blocks):
        """主线程应用渲染结果"""
        try:
            # 预览区默认只读：渲染时临时解锁写入，写完再恢复只读
            self.text.configure(state='normal')
            self.text.delete('1.0', 'end')
            
            # 清除旧的公式图片，重置计数器
            self.math_images = []
            self.equation_counter = 0
            self._math_token_counter = 0
            self._math_token_map = {}
            
            for block in blocks:
                self._render_block(block)

            if bool(getattr(self, '_readonly', False)):
                self.text.configure(state='disabled')
        except Exception as e:
            print(f"Apply render error: {e}")
        finally:
            self._is_rendering = False
            # 渲染完成后触发一次滚动同步
            if hasattr(self, 'on_scroll') and self.on_scroll:
                # 触发一个微小的滚动变化来强制刷新同步位
                pass

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
                    # 插入隐藏 token，确保回写 Markdown 时不会丢公式
                    try:
                        token = self._new_math_token(formula_text, inline=True)
                        self.text.insert('end', token, ('math_token',))
                    except Exception:
                        pass
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
            # 应用缩放比例
            scale = getattr(self, '_scale', 1.0)
            base_fontsize = 12 if is_inline else 14  # 与正文 16pt 匹配
            render_fontsize = int(base_fontsize * scale)

            formula = (latex or '').strip()
            # 兼容 $...$/$$...$$ 包裹：先剥离，避免变成 $$...$$ 导致解析失败
            if formula.startswith('$$') and formula.endswith('$$') and len(formula) >= 4:
                formula = formula[2:-2].strip()
            elif formula.startswith('$') and formula.endswith('$') and len(formula) >= 2:
                formula = formula[1:-1].strip()

            # 兼容示例中双反斜杠（\int），渲染时归一为 \int
            try:
                formula = formula.replace('\\\\', '\\')
            except Exception:
                pass

            if not formula:
                return None

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
                dpi=120,  # 固定 DPI，缩放只通过 fontsize 控制
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
            # 插入隐藏 token，确保回写 Markdown 时不会丢公式
            try:
                token = self._new_math_token(formula_text, inline=False)
                self.text.insert('end', token, ('math_token',))
            except Exception:
                pass
            self.text.insert('end', f'    ({self.equation_counter})\n')
        else:
            # 直接插入文本，使用 math_block 标签居中
            display = f'    {formula_text}    ({self.equation_counter})'
            self.text.insert('end', display, ('math_block',))
            self.text.insert('end', '\n')
        
        self.text.insert('end', '\n')
    
    def _insert_table(self, table_text: str):
        """插入表格 - 在预览中渲染为网格表，支持格式化内容，居中显示"""
        headers, rows, alignments = parse_table(table_text)
        if not headers:
            self.text.insert('end', table_text + '\n\n', ('body',))
            return

        # 插入换行和居中标记
        self.text.insert('end', '\n')
        
        # 记录表格开始位置
        table_start = self.text.index('end-1c')
        
        # 创建表格容器 Frame，使用 place 实现真正居中
        table_container = tk.Frame(self.text, bg=self.text.cget('bg'))
        
        # 创建表格 Frame
        table_frame = tk.Frame(table_container, bg='#E5E7EB', bd=0)
        table_frame.pack(anchor='center', padx=2, pady=2)

        all_rows = [headers] + rows
        num_cols = len(headers)
        
        # 计算每列最大宽度
        col_widths = []
        for c in range(num_cols):
            max_width = 8  # 最小宽度
            for row in all_rows:
                if c < len(row):
                    cell_len = len(row[c])
                    max_width = max(max_width, min(cell_len + 2, 20))  # 限制最大宽度
            col_widths.append(max_width)

        for r, row in enumerate(all_rows):
            for c in range(num_cols):
                cell_text = row[c] if c < len(row) else ''
                is_header = (r == 0)
                
                # 获取对齐方式
                align = 'center'
                if c < len(alignments):
                    if alignments[c] == 'left':
                        align = 'w'
                    elif alignments[c] == 'right':
                        align = 'e'
                    else:
                        align = 'center'
                
                # 使用 Label 组件简化表格单元格
                cell = tk.Label(
                    table_frame,
                    text=cell_text,
                    font=('黑体' if is_header else '宋体', int(12 * getattr(self, '_scale', 1.0)), 'bold' if is_header else 'normal'),
                    fg='#1E293B',
                    bg='#F1F5F9' if is_header else '#FFFFFF',
                    bd=1,
                    relief='solid',
                    padx=int(8 * getattr(self, '_scale', 1.0)),
                    pady=4,
                    width=int(col_widths[c] * getattr(self, '_scale', 1.0)),
                    anchor=align if align != 'center' else 'center',
                )
                cell.grid(row=r, column=c, sticky='nsew', padx=0, pady=0)

        # 配置列权重使表格均匀分布
        for c in range(num_cols):
            table_frame.grid_columnconfigure(c, weight=1)

        # 插入表格到 Text 中
        self.text.window_create('end', window=table_container)
        
        # 应用居中标签
        table_end = self.text.index('end-1c')
        self.text.tag_add('table_block', table_start, table_end)
        
        self.text.insert('end', '\n\n')

        # 插入隐藏的表格 Markdown，确保回写 Markdown 时不会丢表格
        try:
            raw = (table_text or '').strip('\n')
            if raw:
                self.text.insert('end', raw + '\n\n', ('math_token',))
        except Exception:
            pass
    
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
            indent = '    ' * item_level
            
            if is_task:
                # 避免 ☑/☐ 在部分字体下显示为方块导致“乱码/错位”
                checkbox = '☑' if checked else '☐'
                self.text.insert('end', f'{indent}{checkbox} ', ('list_item',))
                self._insert_inline_elements(item_text, extra_tags=['list_item'])
            else:
                if ordered:
                    number = get_number_format(item_level, level_counters[item_level])
                    self.text.insert('end', f'{indent}{number} ', ('list_item',))
                else:
                    # 避免 •/◦ 等符号在部分字体下显示异常
                    self.text.insert('end', f'{indent}• ', ('list_item',))
                self._insert_inline_elements(item_text, extra_tags=['list_item'])
            
            self.text.insert('end', '\n')
        self.text.insert('end', '\n')
    
    def _insert_image(self, alt: str, url: str):
        """插入图片占位"""
        self.text.insert('end', f'🖼️ [{alt}]\n\n')

    # ==================== 点击跳转方法 ====================
    
    def _on_double_click(self, event):
        """双击预览区跳转到对应的源码行"""
        try:
            # 获取点击位置
            index = self.text.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0])
            
            # 查找对应的源码行
            source_line = self._find_source_line_for_preview_line(line_num)
            
            if source_line and self.on_jump_to_line:
                self.on_jump_to_line(source_line)
                
                # 高亮显示跳转的行（短暂闪烁效果）
                self._highlight_preview_line(line_num)
        except Exception:
            pass
    
    def _find_source_line_for_preview_line(self, preview_line: int) -> int:
        """根据预览区行号查找对应的源码行号"""
        # 从 paragraph_map 中查找最近的映射
        if not self.paragraph_map:
            return preview_line  # 回退到简单映射
        
        # 查找最近的段落映射
        nearest_line = None
        min_distance = float('inf')
        
        for line_start, info in self.paragraph_map.items():
            try:
                # line_start 是预览区的行位置
                line_pos = int(str(line_start).split('.')[0])
                distance = abs(line_pos - preview_line)
                if distance < min_distance:
                    min_distance = distance
                    nearest_line = info.get('md_line', preview_line)
            except Exception:
                continue
        
        return nearest_line if nearest_line else preview_line
    
    def _highlight_preview_line(self, line_num: int):
        """短暂高亮预览区的行"""
        try:
            # 添加高亮标签
            self.text.tag_configure('jump_highlight', background='#fef3c7')
            self.text.tag_add('jump_highlight', f"{line_num}.0", f"{line_num}.end")
            
            # 300ms 后移除高亮
            self.after(300, lambda: self._remove_highlight(line_num))
        except Exception:
            pass
    
    def _remove_highlight(self, line_num: int):
        """移除行高亮"""
        try:
            self.text.tag_remove('jump_highlight', f"{line_num}.0", f"{line_num}.end")
        except Exception:
            pass
    
    def set_jump_callback(self, callback):
        """设置跳转回调函数
        
        Args:
            callback: 回调函数，接收源码行号作为参数
        """
        self.on_jump_to_line = callback

    def _jump_to_editor_from_context(self):
        """右键菜单触发跳转：定位到当前预览行对应的源码"""
        try:
            index = self.text.index(tk.INSERT)
            line_num = int(index.split('.')[0])
            source_line = self._find_source_line_for_preview_line(line_num)
            if source_line and self.on_jump_to_line:
                self.on_jump_to_line(source_line)
                self._highlight_preview_line(line_num)
        except Exception:
            pass

    # ==================== 滚动同步方法 ====================
    
    def _on_scrollbar(self, *args):
        """滚动条事件处理"""
        self.text.yview(*args)
        # 触发滚动同步
        try:
            first = self.text.yview()[0]
            if (
                getattr(self, '_sync_scroll_enabled', True)
                and hasattr(self, 'on_scroll')
                and self.on_scroll
                and not getattr(self, '_scroll_updating', False)
            ):
                self.on_scroll(float(first))
        except Exception:
            pass
    
    def _on_mousewheel(self, event):
        """鼠标滚轮事件处理"""
        # 执行滚动
        if event.num == 4:  # Linux scroll up
            self.text.yview_scroll(-3, "units")
        elif event.num == 5:  # Linux scroll down
            self.text.yview_scroll(3, "units")
        else:  # Windows/Mac
            self.text.yview_scroll(-1 * (event.delta // 120), "units")
        
        # 触发滚动同步
        try:
            first = self.text.yview()[0]
            if hasattr(self, 'on_scroll') and self.on_scroll and not getattr(self, '_scroll_updating', False):
                self.on_scroll(float(first))
        except Exception:
            pass
        
        return "break"
    
    def _on_text_scroll(self, first, last):
        """文本滚动事件处理，同步到编辑器"""
        # 更新滚动条
        self.scrollbar.set(first, last)
        
        # 触发滚动同步回调
        if (
            getattr(self, '_sync_scroll_enabled', True)
            and hasattr(self, 'on_scroll')
            and self.on_scroll
            and not getattr(self, '_scroll_updating', False)
        ):
            try:
                self.on_scroll(float(first))
            except Exception:
                pass
    
    def sync_scroll_to(self, position: float):
        """同步滚动到指定位置
        
        Args:
            position: 滚动位置 (0.0 - 1.0)
        """
        if getattr(self, '_scroll_updating', False):
            return
        
        self._scroll_updating = True
        try:
            self.text.yview_moveto(position)
        except Exception:
            pass
        finally:
            self._scroll_updating = False
    
    def set_sync_scroll_enabled(self, enabled: bool):
        """设置是否启用同步滚动
        
        Args:
            enabled: 是否启用
        """
        self._sync_scroll_enabled = enabled
    
    def apply_theme(self, theme_config: dict):
        """应用预览主题
        
        Args:
            theme_config: 主题配置字典，来自 PreviewTheme.to_tkinter_config()
        """
        try:
            # 更新背景色
            bg = theme_config.get('background', '#FFFFFF')
            fg = theme_config.get('text_color', '#111827')
            self.text.configure(bg=bg, fg=fg)
            
            # 更新字体
            font_family = theme_config.get('font_family', '宋体')
            font_size = theme_config.get('font_size', 16)
            self.text.configure(font=(font_family, font_size))
            
            # 更新标签样式
            h1_color = theme_config.get('h1_color', '#1f2937')
            h1_size = theme_config.get('h1_size', 28)
            self.text.tag_configure('h1', foreground=h1_color, font=('黑体', h1_size, 'bold'))
            
            h2_color = theme_config.get('h2_color', '#374151')
            h2_size = theme_config.get('h2_size', 22)
            self.text.tag_configure('h2', foreground=h2_color, font=('黑体', h2_size, 'bold'))
            
            h3_color = theme_config.get('h3_color', '#4b5563')
            h3_size = theme_config.get('h3_size', 18)
            self.text.tag_configure('h3', foreground=h3_color, font=('黑体', h3_size, 'bold'))
            
            h4_color = theme_config.get('h4_color', '#6b7280')
            h4_size = theme_config.get('h4_size', 16)
            self.text.tag_configure('h4', foreground=h4_color, font=('黑体', h4_size, 'bold'))
            
            # 链接颜色
            link_color = theme_config.get('link_color', '#0000FF')
            self.text.tag_configure('link', foreground=link_color)
            
            # 代码样式
            code_bg = theme_config.get('code_bg', '#F5F5F5')
            code_color = theme_config.get('code_color', '#1F2937')
            code_font = theme_config.get('code_font', 'Consolas')
            self.text.tag_configure('code', background=code_bg, foreground=code_color, font=(code_font, 10))
            
            code_block_bg = theme_config.get('code_block_bg', '#FAFAFA')
            code_block_color = theme_config.get('code_block_color', '#1F2937')
            self.text.tag_configure('code_block', background=code_block_bg, foreground=code_block_color, font=(code_font, 10))
            
            # 引用样式
            blockquote_bg = theme_config.get('blockquote_bg', '#f9fafb')
            blockquote_color = theme_config.get('blockquote_color', '#6B7280')
            self.text.tag_configure('quote', background=blockquote_bg, foreground=blockquote_color)
            
        except Exception as e:
            print(f"应用预览主题失败: {e}")
