# -*- coding: utf-8 -*-

from io import BytesIO
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
import re

from parser import parse_markdown, parse_inline, parse_table, InlineType
from utils import normalize_markdown, convert_latex_delimiters
from ui.theme import COLORS


class MarkdownPreview(ctk.CTkFrame):
    """可编辑的Markdown预览组件 - 支持双向同步和滚动同步"""
    def __init__(self, master, on_content_change=None, app=None, on_scroll=None, **kwargs):
        super().__init__(master, fg_color=COLORS['bg_card'], corner_radius=12, **kwargs)
        self._base_width = 900
        self._resize_timer = None
        self._pending_scroll_ratio = None
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
        self._performance_mode = "normal"
        self._render_batch_size = 24
        self._render_batch_delay_ms = 1
        self._render_token = 0
        self._pending_markdown_text = None
        self._last_requested_markdown = None
        self._last_rendered_markdown = None
        self._last_parsed_markdown = None
        self._last_parsed_blocks = None
        
        # 是否正在更新（防止循环触发）
        self._updating = False
        
        # 存储段落信息：{line_start: {'type': 'paragraph', 'md_line': 1, 'format': []}}
        self.paragraph_map = {}
        
        # 脚注存储
        self._footnotes = {} # ref -> content

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
        
        # 浮动大纲 (ToC)
        self._toc_visible = False
        self._toc_frame = None
        self._toc_listbox = None
        self._headings = [] # [(level, title, source_line)]
        
        # 平滑滚动相关
        self._smooth_scroll_timer = None
        
        # 阅读进度条
        self._progress_bar = None
        self._init_reading_progress_bar()
        
        # 右键菜单
        self._create_context_menu()
        
        # 监听尺寸变化，驱动平滑缩放
        self.bind("<Configure>", self._on_resize)
    
    def _init_reading_progress_bar(self):
        """初始化预览区顶部的阅读进度条"""
        self._progress_canvas = tk.Canvas(
            self,
            height=3,
            bg=self._get_bg_color(),
            highlightthickness=0,
            bd=0
        )
        self._progress_canvas.place(relx=0, rely=0, relwidth=1, y=0)
        
        # 进度条线条
        self._progress_line = self._progress_canvas.create_line(
            0, 1, 0, 1, 
            fill=COLORS.get('primary', '#3b82f6'), 
            width=3
        )

    def _update_reading_progress(self, position: float):
        """更新进度条长度"""
        if not hasattr(self, '_progress_canvas'): return
        
        # position 是 0.0 - 1.0 的滚动比例
        # 计算可见范围
        yview = self.text.yview()
        # 实际进度应该是 可见区域底部所占的比例
        # 如果到底了就是 100%
        progress = yview[1]
        
        canvas_width = self._progress_canvas.winfo_width()
        if canvas_width <= 1: 
            canvas_width = self.winfo_width()
            
        self._progress_canvas.coords(self._progress_line, 0, 1, int(progress * canvas_width), 1)
    
    def _get_bg_color(self):
        """兼容 CTkFrame，获取背景色"""
        try:
            bg = self.cget("fg_color")
            if isinstance(bg, (list, tuple)):
                return bg[0]
            return bg
        except Exception:
            return "#FFFFFF"
    
    def _on_resize(self, event):
        """窗口尺寸变化时，动态调整缩放比例并平滑重渲染"""
        try:
            new_scale = max(0.8, min(1.6, event.width / float(self._base_width)))
            min_delta = 0.08 if self._performance_mode == "high" else 0.03
            if abs(new_scale - self._scale) < min_delta:
                return
            self._scale = new_scale
            # 防抖重渲染
            if self._resize_timer:
                self.after_cancel(self._resize_timer)
            delay = 220 if self._performance_mode == "high" else 120
            self._resize_timer = self.after(delay, self._rerender_after_resize)
        except Exception:
            pass
    
    def _rerender_after_resize(self):
        """根据当前缩放重新渲染，确保文字与表格等比例放大"""
        self._resize_timer = None
        try:
            if self.app and hasattr(self.app, 'input_text'):
                md = self.app.input_text.get("1.0", "end-1c")
                self.update_preview(md)
        except Exception:
            pass

    def _on_text_scroll(self, first, last):
        """滚动事件处理"""
        self._update_reading_progress(float(last))
        if self.on_scroll:
            self.on_scroll(first, last)

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
            self.text.pack(side="left", fill="both", expand=True, padx=30, pady=20)
            # 去掉固定宽度，采用边框模拟纸张
            self.text.configure(width=0, borderwidth=1, relief="solid")
            self.scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=20)
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
            # 同时使用 elide 和极小字号 + 背景色隐藏，确保在各种环境下都不留痕迹
            self.text.tag_configure('math_token', 
                elide=True, 
                font=('TkDefaultFont', 1), 
                foreground=self.text.cget('bg')
            )
        except Exception:
            try:
                self.text.tag_configure('math_token', 
                    font=('TkDefaultFont', 1),
                    foreground=self.text.cget('bg')
                )
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
        # 优先级：Cambria Math (Office 原生), STIX Two Math, Times New Roman (通用), 
        # Linux Libertine, DejaVu Serif (常用替代)
        math_fonts = [
            'Cambria Math', 'STIX Two Math', 'Latin Modern Math', 
            'Cambria', 'Times New Roman', 'Liberation Serif',
            'DejaVu Serif', 'SimSun', 'Microsoft YaHei'
        ]
        
        for f in math_fonts:
            if f in available:
                return f
        return 'serif'
    
    def update_preview(self, markdown_text: str):
        """?????? - ???"""
        import threading

        if markdown_text == self._last_rendered_markdown and not getattr(self, '_is_rendering', False):
            return

        self._last_requested_markdown = markdown_text

        if getattr(self, '_is_rendering', False):
            self._pending_markdown_text = markdown_text
            return

        self._render_token += 1
        render_token = self._render_token
        self._pending_markdown_text = None

        self._start_render_animation()

        def _render_task():
            self._is_rendering = True
            try:
                if markdown_text == self._last_parsed_markdown and self._last_parsed_blocks is not None:
                    blocks = self._last_parsed_blocks
                else:
                    processed_text = convert_latex_delimiters(markdown_text)
                    processed_text = normalize_markdown(processed_text)
                    blocks = parse_markdown(processed_text)
                    self._last_parsed_markdown = markdown_text
                    self._last_parsed_blocks = blocks

                self.after(0, lambda: self._apply_render_result(render_token, blocks, markdown_text))
            except Exception as e:
                print(f"Render error: {e}")
                self.after(0, lambda: self._set_preview_error())
                self.after(0, lambda: self._finish_render(None))

        threading.Thread(target=_render_task, daemon=True).start()

    def _start_render_animation(self):
        """在进度条位置显示渲染中的呼吸动画"""
        if not hasattr(self, '_progress_canvas'): return
        self._progress_canvas.configure(bg=COLORS.get('highlight', '#f0f9ff'))
        self._step_render_animation(0)

    def _step_render_animation(self, step):
        """渲染动画单步"""
        if not getattr(self, '_is_rendering', False):
            self._progress_canvas.configure(bg=self._get_bg_color())
            # 渲染结束，恢复正常的进度条显示
            self._update_reading_progress(0) 
            return
            
        # 简单的颜色交替
        colors = ["#3b82f6", "#60a5fa", "#93c5fd", "#60a5fa"]
        color = colors[step % len(colors)]
        self._progress_canvas.itemconfig(self._progress_line, fill=color)
        
        self.after(200, lambda: self._step_render_animation(step + 1))

    def _apply_render_result(self, render_token, blocks, markdown_text):
        """?????????"""
        try:
            if render_token != self._render_token:
                self._finish_render(None)
                return

            try:
                self._pending_scroll_ratio = self.text.yview()[0]
            except Exception:
                self._pending_scroll_ratio = None

            self.text.configure(state='normal')
            self.text.delete('1.0', 'end')

            self.math_images = []
            self.equation_counter = 0
            self._math_token_counter = 0
            self._math_token_map = {}
            self._headings = []
            self._footnotes = {}
            self._anchor_map = {}

            for block in blocks:
                if block.type == 'heading':
                    anchor_id = re.sub(r'[^\w-]', '', block.content.lower().replace(' ', '-'))
                    self._headings.append((block.level, block.content, block.line_start, anchor_id))
                elif block.type == 'footnote_def':
                    self._footnotes[block.language] = block.content

            render_blocks = [block for block in blocks if block.type != 'footnote_def']
            self._render_blocks_in_batches(render_token, render_blocks, 0, markdown_text)
        except Exception as e:
            print(f"Apply render error: {e}")
            self._set_preview_error()
            self._finish_render(None)

    def _render_blocks_in_batches(self, render_token, blocks, start_index, markdown_text):
        try:
            if render_token != self._render_token:
                self._finish_render(None)
                return

            end_index = min(start_index + self._render_batch_size, len(blocks))
            for block in blocks[start_index:end_index]:
                self._render_block(block)

            if end_index < len(blocks):
                self.after(
                    self._render_batch_delay_ms,
                    lambda: self._render_blocks_in_batches(render_token, blocks, end_index, markdown_text),
                )
                return

            self._finalize_render(markdown_text)
        except Exception as e:
            print(f"Batch render error: {e}")
            self._set_preview_error()
            self._finish_render(None)

    def _finalize_render(self, markdown_text: str):
        try:
            if self._footnotes:
                self.text.insert('end', '\n' + '─' * 20 + '\n', ('body',))
                for ref, content in self._footnotes.items():
                    self.text.insert('end', f'[^{ref}]: ', ('supsub',))
                    self._insert_inline_elements(content, extra_tags=['body'])
                    self.text.insert('end', '\n')

            if bool(getattr(self, '_readonly', False)):
                self.text.configure(state='disabled')

            try:
                if self._pending_scroll_ratio is not None:
                    self.text.yview_moveto(self._pending_scroll_ratio)
            except Exception:
                pass
            finally:
                self._pending_scroll_ratio = None

            if self._toc_visible:
                self._update_floating_toc_content()

            self._last_rendered_markdown = markdown_text
        finally:
            self._finish_render(markdown_text)

    def _finish_render(self, markdown_text):
        self._is_rendering = False
        pending = self._pending_markdown_text
        self._pending_markdown_text = None
        if pending and pending != markdown_text:
            self.after(0, lambda: self.update_preview(pending))

    def set_performance_mode(self, mode: str):
        mode = 'high' if str(mode).lower() == 'high' else 'normal'
        self._performance_mode = mode
        if mode == 'high':
            self._render_batch_size = 6
            self._render_batch_delay_ms = 2
        else:
            self._render_batch_size = 24
            self._render_batch_delay_ms = 1

    def _render_block(self, block):
        """渲染块级元素"""
        if block.type == 'heading':
            self._insert_heading(block.content, block.level)
        
        elif block.type == 'paragraph':
            self._insert_paragraph(block.content)
        
        elif block.type == 'code_block':
            self._insert_code_block(block.content, block.language)
        
        elif block.type == 'math_block':
            self._insert_math_block(block.content, env_name=block.language)
        
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
        """插入标题 - 使用共用解析器处理行内元素，并记录锚点位置"""
        tag = f'h{min(level, 4)}'
        anchor_id = re.sub(r'[^\w-]', '', text.lower().replace(' ', '-'))
        
        # 记录锚点位置
        start_idx = self.text.index('end-1c')
        self._anchor_map[anchor_id] = start_idx
        
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
                # 检查是否是内部锚点链接
                if elem.url.startswith('#'):
                    anchor_id = elem.url[1:]
                    tag_name = f"anchor_link_{anchor_id}"
                    self.text.insert('end', elem.content, tuple(tags + ['link', tag_name]))
                    self.text.tag_bind(tag_name, '<Button-1>', lambda e, aid=anchor_id: self._jump_to_anchor(aid))
                else:
                    self.text.insert('end', elem.content, tuple(tags + ['link']))
            
            elif elem.type == InlineType.IMAGE:
                self.text.insert('end', f'🖼️[{elem.content}]', tuple(tags))
            
            elif elem.type == InlineType.STRIKETHROUGH:
                self.text.insert('end', elem.content, tuple(tags + ['strikethrough']))
            
            elif elem.type == InlineType.SUPERSCRIPT:
                self.text.insert('end', elem.content, tuple(tags + ['superscript']))
            
            elif elem.type == InlineType.SUBSCRIPT:
                self.text.insert('end', elem.content, tuple(tags + ['subscript']))
            
            elif elem.type == InlineType.FOOTNOTE_REF:
                # 脚注引用 [^1]
                ref = elem.content
                tag_name = f"fn_ref_{ref}"
                self.text.insert('end', f'[^{ref}]', tuple(tags + ['superscript', 'link', tag_name]))
                # 绑定悬浮预览
                self.text.tag_bind(tag_name, '<Enter>', lambda e, r=ref: self._show_footnote_tooltip(e, r))
                self.text.tag_bind(tag_name, '<Leave>', lambda e: self._hide_footnote_tooltip())
    
    def _insert_code_block(self, code: str, language: str = ''):
        """插入代码块 - 支持 Mermaid/PlantUML 渲染"""
        lang = (language or '').lower().strip()
        
        # 1. 检查是否是图表语言
        if lang in ['mermaid', 'plantuml']:
            # 尝试实时渲染图表
            img = self._render_diagram_to_img(code, lang)
            if img:
                self.text.insert('end', f'[{lang} 图表]\n', ('code',))
                self.math_images.append(img)
                self.text.image_create('end', image=img)
                self.text.insert('end', '\n\n')
                return

        # 2. 普通代码块渲染
        if language:
            self.text.insert('end', f'[{language}]\n', ('code',))
        
        # 记录起始位置用于添加复制按钮
        code_start = self.text.index('end-1c')
        self.text.insert('end', code + '\n\n', ('code_block',))
        code_end = self.text.index('end-1c')
        
        # 为代码块添加简单的“复制”按钮图标
        try:
            copy_btn = tk.Label(self.text, text="📋", font=('Segoe UI Emoji', 10), cursor='hand2', bg='#FAFAFA', fg='#6B7280')
            copy_btn.bind('<Button-1>', lambda e, c=code: self._copy_to_clipboard(c))
            self.text.window_create(code_start, window=copy_btn)
        except Exception:
            pass

    def _render_diagram_to_img(self, code: str, lang: str):
        """调用 DiagramFeature 渲染图表"""
        try:
            if not self.app or not hasattr(self.app, 'diagram_feature'):
                return None
            
            # 使用临时文件存储渲染结果
            output_path = os.path.join(tempfile.gettempdir(), f"diag_{hash(code)}.png")
            
            success = False
            if lang == 'mermaid':
                success = self.app.diagram_feature.mermaid_renderer.render(code, output_path)
            elif lang == 'plantuml':
                success = self.app.diagram_feature.plantuml_renderer.render(code, output_path)
                
            if success and os.path.exists(output_path):
                img = Image.open(output_path)
                # 限制最大宽度
                max_w = self.text.winfo_width() - 100
                if max_w < 200: max_w = 600
                if img.width > max_w:
                    ratio = max_w / img.width
                    img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
                
                return ImageTk.PhotoImage(img)
        except Exception:
            pass
        return None

    def _copy_to_clipboard(self, text: str):
        """内部复制方法"""
        self.clipboard_clear()
        self.clipboard_append(text)
        if self.app:
            self.app.update_status("✅ 代码已复制到剪贴板")
    
    def _render_latex(self, latex: str, fontsize: int = 16, is_inline: bool = False) -> ImageTk.PhotoImage:
        """使用 matplotlib.mathtext 渲染 LaTeX 公式为图片。"""
        try:
            plt.rcParams['mathtext.fontset'] = 'cm'
            formula = (latex or '').strip()
            
            # 1. 处理常见的不受 matplotlib 支持的环境块
            if '\\begin{' in formula:
                # 特殊处理 align 环境：将其转换为单行显示
                if 'align' in formula:
                    formula = re.sub(r'\\begin\{align\*?\}', '', formula)
                    formula = re.sub(r'\\end\{align\*?\}', '', formula)
                    formula = formula.replace('&', ' ')
                    formula = formula.replace('\\\\', ' \\quad ')
                else:
                    formula = re.sub(r'\\begin\{[a-z*]+\}', '', formula)
                    formula = re.sub(r'\\end\{[a-z*]+\}', '', formula)
                    formula = formula.replace('&', ' ')
            
            # 2. 清理多余的 $ 符号
            if formula.startswith('$$') and formula.endswith('$$'):
                formula = formula[2:-2].strip()
            elif formula.startswith('$') and formula.endswith('$'):
                formula = formula[1:-1].strip()
            
            if not formula:
                return None

            fig, ax = plt.subplots(figsize=(0.01, 0.01))
            ax.axis('off')
            scale = getattr(self, '_scale', 1.0)
            base_fontsize = 12 if is_inline else 14
            render_fontsize = int(base_fontsize * scale)

            text = ax.text(0.5, 0.5, f'${formula}$', fontsize=render_fontsize,
                          ha='center', va='center', transform=ax.transAxes, color='#1a1a2e')

            fig.canvas.draw()
            bbox = text.get_window_extent(fig.canvas.get_renderer())
            bbox = bbox.expanded(1.1, 1.2)
            fig.set_size_inches(bbox.width / fig.dpi, bbox.height / fig.dpi)

            buf = BytesIO()
            fig.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.05, dpi=120, transparent=True)
            plt.close(fig)
            buf.seek(0)
            return ImageTk.PhotoImage(Image.open(buf))
        except Exception:
            return None

    def _insert_math_block(self, formula: str, env_name: str = ""):
        """插入块级公式：居中显示，右侧编号"""
        self.equation_counter += 1
        formula_text = formula.strip()
        
        if not formula_text:
            return
        
        # 预清理：剥离多余的定界符
        if formula_text.startswith('$$') and formula_text.endswith('$$'):
            formula_text = formula_text[2:-2].strip()
        elif formula_text.startswith('$') and formula_text.endswith('$'):
            formula_text = formula_text[1:-1].strip()
        elif formula_text.startswith('\\[') and formula_text.endswith('\\]'):
            formula_text = formula_text[2:-2].strip()

        # 尝试渲染公式为图片
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
            # 渲染失败回退：针对 align 环境进行视觉优化
            self.text.insert('end', '          ', ('math_block',))
            
            display_text = formula_text
            # 如果是特定的数学环境，进行格式化
            if env_name or '\\begin{' in display_text:
                # 仅移除外层包装标签
                display_text = re.sub(r'\\begin\{[a-z*]+\}', '', display_text)
                display_text = re.sub(r'\\end\{[a-z*]+\}', '', display_text)
                # 处理对齐符号 & 替换为空格
                display_text = display_text.replace('&', ' ')
                # 处理 \\ 换行符，确保在 Text 组件中产生真实换行，并应用居中标签
                lines = display_text.split('\\\\')
                for idx, line in enumerate(lines):
                    line_content = line.strip()
                    if not line_content: continue
                    self.text.insert('end', line_content, ('math', 'math_block'))
                    if idx < len(lines) - 1:
                        self.text.insert('end', '\n          ', ('math_block',))
            else:
                self.text.insert('end', display_text.strip(), ('math', 'math_block'))
            
            self.text.insert('end', f'    ({self.equation_counter})\n', ('math_block',))
        
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
                # 为单元格绑定表格工具菜单，传入单元格内容和整行内容
                cell.bind('<Button-3>', lambda e, t=cell_text, row=row, full=table_text: self._show_table_menu(e, t, row, full))

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
                tag_name = f"task_cb_{item.get('line', 0)}"
                self.text.insert('end', f'{indent}{checkbox}', ('list_item', 'task_checkbox', tag_name))
                self.text.insert('end', ' ', ('list_item',))
                self._insert_inline_elements(item_text, extra_tags=['list_item'])
                
                # 为复选框添加点击绑定（如果尚未绑定）
                self.text.tag_bind('task_checkbox', '<Button-1>', self._on_checkbox_click)
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
        # 如果是本地图片或支持加载的，可以尝试显示，这里先简单实现点击查看
        img_label = tk.Label(self.text, text=f'🖼️ [{alt}]', font=('Microsoft YaHei', 10), cursor='hand2', fg=COLORS.get('primary', '#3b82f6'), bg=self.text.cget('bg'))
        img_label.bind('<Button-1>', lambda e: self._zoom_image(url))
        self.text.window_create('end', window=img_label)
        self.text.insert('end', '\n\n')

    def _zoom_image(self, url: str):
        """查看图片大图"""
        if not url: return
        try:
            # 如果是相对路径，转换为绝对路径
            if not os.path.isabs(url) and self.app and self.app.current_file:
                url = os.path.join(os.path.dirname(self.app.current_file), url)
            
            if os.path.exists(url):
                # 简单实现：使用系统默认应用打开
                import platform
                import subprocess
                if platform.system() == 'Windows':
                    os.startfile(url)
                elif platform.system() == 'Darwin':
                    subprocess.run(['open', url])
                else:
                    subprocess.run(['xdg-open', url])
        except Exception as e:
            print(f"Zoom image error: {e}")

    def _show_table_menu(self, event, cell_text, row_data, table_text):
        """显示表格快捷菜单"""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="复制单元格", command=lambda: self._copy_to_clipboard(cell_text))
        menu.add_command(label="复制整行 (CSV)", command=lambda: self._copy_row_as_csv(row_data))
        menu.add_separator()
        menu.add_command(label="搜索并过滤此列", command=lambda: self._filter_table_by_cell(table_text, cell_text))
        menu.add_command(label="转置表格预览", command=lambda: self._transpose_table(table_text))
        menu.add_separator()
        menu.add_command(label="整个表格复制为 CSV", command=lambda: self._copy_table_as(table_text, 'csv'))
        menu.add_command(label="整个表格复制为 Markdown", command=lambda: self._copy_table_as(table_text, 'md'))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _filter_table_by_cell(self, table_text, filter_val):
        """简单的表格过滤功能：根据选中内容过滤行"""
        # 这个功能通常需要重新渲染表格，这里我们通过复制过滤后的 MD 到剪贴板或弹窗实现
        try:
            headers, rows, _ = parse_table(table_text)
            if not headers: return
            
            # 过滤包含该值的行
            filtered_rows = [row for row in rows if any(filter_val in str(cell) for cell in row)]
            
            # 构建新的 Markdown
            new_md = []
            new_md.append("| " + " | ".join(headers) + " |")
            new_md.append("| " + " | ".join(["---"] * len(headers)) + " |")
            for row in filtered_rows:
                new_md.append("| " + " | ".join(row) + " |")
            
            filtered_text = "\n".join(new_md)
            self._copy_to_clipboard(filtered_text)
            if self.app:
                self.app.update_status(f"🔍 已过滤表格并复制 (剩余 {len(filtered_rows)} 行)", is_temp=True)
        except Exception as e:
            print(f"Filter table error: {e}")

    def _transpose_table(self, table_text):
        """转置表格内容并重新渲染（临时预览）"""
        try:
            headers, rows, alignments = parse_table(table_text)
            if not headers: return
            
            all_data = [headers] + rows
            # 使用 zip 转置
            transposed = list(map(list, zip(*all_data)))
            
            # 构建新的 Markdown 表格字符串
            new_md = []
            new_md.append("| " + " | ".join(transposed[0]) + " |")
            new_md.append("| " + " | ".join(["---"] * len(transposed[0])) + " |")
            for row in transposed[1:]:
                new_md.append("| " + " | ".join(row) + " |")
            
            transposed_text = "\n".join(new_md)
            
            # 这里简单处理：弹窗显示或提示用户复制。
            # 更好的做法是原地替换预览，但因为预览是全量生成的，这比较复杂。
            # 先实现复制到剪贴板并提示。
            self._copy_to_clipboard(transposed_text)
            if self.app:
                self.app.update_status("🔄 已转置表格并复制到剪贴板", is_temp=True)
        except Exception as e:
            print(f"Transpose table error: {e}")

    def _copy_row_as_csv(self, row_data):
        """复制整行数据为 CSV 格式"""
        import csv
        import io
        si = io.StringIO()
        cw = csv.writer(si)
        cw.writerow(row_data)
        self._copy_to_clipboard(si.getvalue().strip())

    def _copy_table_as(self, table_text, fmt):
        """将表格转换格式并复制到剪贴板"""
        try:
            if fmt == 'md':
                content = table_text
            else:
                headers, rows, _ = parse_table(table_text)
                import csv
                import io
                output = BytesIO() # Use BytesIO or io.StringIO
                si = io.StringIO()
                cw = csv.writer(si)
                cw.writerow(headers)
                cw.writerows(rows)
                content = si.getvalue()
            
            self.clipboard_clear()
            self.clipboard_append(content)
            if self.app:
                self.app.update_status(f"✅ 表格已复制为 {fmt.upper()}")
        except Exception as e:
            print(f"Copy table error: {e}")

    def _show_footnote_tooltip(self, event, ref):
        """显示脚注内容悬浮窗"""
        content = self._footnotes.get(ref)
        if not content: return
        
        # 简单实现：使用 Toplevel
        self._fn_tooltip = tk.Toplevel(self)
        self._fn_tooltip.wm_overrideredirect(True)
        self._fn_tooltip.wm_geometry(f"+{event.x_root+15}+{event.y_root+10}")
        
        frame = tk.Frame(self._fn_tooltip, bg='#ffffe1', padx=5, pady=5, borderwidth=1, relief='solid')
        frame.pack()
        
        lbl = tk.Label(frame, text=content, bg='#ffffe1', font=('Microsoft YaHei', 9), wraplength=300, justify='left')
        lbl.pack()

    def _hide_footnote_tooltip(self):
        """隐藏脚注悬浮窗"""
        if hasattr(self, '_fn_tooltip'):
            self._fn_tooltip.destroy()
            del self._fn_tooltip

    def _on_checkbox_click(self, event):
        """点击预览区复选框，同步修改源码"""
        try:
            index = self.text.index(f"@{event.x},{event.y}")
            tags = self.text.tag_names(index)
            for tag in tags:
                if tag.startswith('task_cb_'):
                    line_num = int(tag.replace('task_cb_', ''))
                    if line_num > 0:
                        self._toggle_source_checkbox(line_num)
                    break
            return "break"
        except Exception:
            pass

    def _toggle_source_checkbox(self, line_num: int):
        """切换源码中对应行的复选框状态"""
        if not self.app or not hasattr(self.app, 'input_text'):
            return
        try:
            text_widget = self.app.input_text._textbox
            line_content = text_widget.get(f"{line_num}.0", f"{line_num}.end")
            new_content = None
            if '[ ]' in line_content:
                new_content = line_content.replace('[ ]', '[x]', 1)
            elif '[x]' in line_content:
                new_content = line_content.replace('[x]', '[ ]', 1)
            elif '[X]' in line_content:
                new_content = line_content.replace('[X]', '[ ]', 1)
            if new_content:
                text_widget.delete(f"{line_num}.0", f"{line_num}.end")
                text_widget.insert(f"{line_num}.0", new_content)
                if hasattr(self.app, 'on_text_change'):
                    self.app.on_text_change(None)
        except Exception:
            pass

    def _on_double_click(self, event):
        """鼠标双击事件处理：双击跳转到源码"""
        try:
            index = self.text.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0])
            source_line = self._find_source_line_for_preview_line(line_num)
            if source_line and self.on_jump_to_line:
                self.on_jump_to_line(source_line)
                self._highlight_preview_line(line_num)
        except Exception:
            pass

    def _find_source_line_for_preview_line(self, preview_line: int) -> int:
        """根据预览区行号查找对应的源码行号"""
        if not self.paragraph_map:
            return preview_line
        nearest_line = None
        min_distance = float('inf')
        for line_start, info in self.paragraph_map.items():
            try:
                line_pos = int(str(line_start).split('.')[0])
                distance = abs(line_pos - preview_line)
                if distance < min_distance:
                    min_distance = distance
                    nearest_line = info.get('md_line', preview_line)
            except Exception:
                continue
        return nearest_line if nearest_line else preview_line

    def toggle_floating_toc(self):
        """切换浮动大纲显示"""
        self._toc_visible = not self._toc_visible
        if self._toc_visible:
            self._show_floating_toc()
        else:
            self._hide_floating_toc()

    def _show_floating_toc(self):
        """显示浮动大纲 - 支持拖拽"""
        if not self._toc_frame:
            self._toc_frame = tk.Frame(
                self, 
                bg=COLORS.get('bg_card', '#ffffff'),
                highlightthickness=1,
                highlightbackground=COLORS.get('border', '#e5e7eb'),
                cursor='fleur' # 移动光标
            )
            
            # 标题栏（用于拖拽）
            title_lbl = tk.Label(
                self._toc_frame, 
                text="文档大纲 ✥", 
                font=('Microsoft YaHei', 9, 'bold'),
                bg=COLORS.get('bg_card', '#ffffff'),
                fg=COLORS.get('text_secondary', '#6b7280'),
                cursor='fleur'
            )
            title_lbl.pack(side='top', fill='x', padx=5, pady=2)
            
            # 绑定拖拽事件
            title_lbl.bind('<Button-1>', self._on_toc_drag_start)
            title_lbl.bind('<B1-Motion>', self._on_toc_drag_motion)
            
            self._toc_listbox = tk.Listbox(
                self._toc_frame,
                font=('Microsoft YaHei', 9),
                bg=COLORS.get('bg_card', '#ffffff'),
                fg=COLORS.get('text_primary', '#111827'),
                borderwidth=0,
                highlightthickness=0,
                selectbackground=COLORS.get('primary', '#3b82f6'),
                selectforeground='white',
                activestyle='none'
            )
            self._toc_listbox.pack(side='top', fill='both', expand=True, padx=2, pady=2)
            self._toc_listbox.bind('<<ListboxSelect>>', self._on_toc_select)
            self._toc_frame.place(relx=1.0, rely=0.1, anchor='ne', width=180, height=300, x=-20)
        self._toc_frame.lift()
        self._update_floating_toc_content()

    def _on_toc_drag_start(self, event):
        """记录拖拽起始位置"""
        self._toc_drag_data = {"x": event.x, "y": event.y}

    def _on_toc_drag_motion(self, event):
        """处理拖拽移动"""
        if not hasattr(self, '_toc_drag_data'): return
        
        delta_x = event.x - self._toc_drag_data["x"]
        delta_y = event.y - self._toc_drag_data["y"]
        
        new_x = self._toc_frame.winfo_x() + delta_x
        new_y = self._toc_frame.winfo_y() + delta_y
        
        # 限制在预览区范围内
        max_x = self.winfo_width() - self._toc_frame.winfo_width()
        max_y = self.winfo_height() - self._toc_frame.winfo_height()
        
        new_x = max(0, min(new_x, max_x))
        new_y = max(0, min(new_y, max_y))
        
        self._toc_frame.place(relx=0, rely=0, anchor='nw', x=new_x, y=new_y)

    def _hide_floating_toc(self):
        """隐藏浮动大纲"""
        if self._toc_frame:
            self._toc_frame.place_forget()
        self._toc_visible = False

    def _update_floating_toc_content(self):
        """更新浮动大纲内容"""
        if not self._toc_listbox:
            return

        # 如果没有记录的标题，尝试从当前文本内容重建，避免空白
        if not self._headings:
            try:
                rebuilt = []
                lines = self.text.get("1.0", "end-1c").split("\n")
                for idx, line in enumerate(lines, start=1):
                    stripped = line.lstrip()
                    if stripped.startswith("#"):
                        level = len(stripped) - len(stripped.lstrip("#"))
                        title = stripped.lstrip("#").strip()
                        if title:
                            rebuilt.append((level, title, idx))
                if rebuilt:
                    self._headings = rebuilt
            except Exception:
                pass

        self._toc_listbox.delete(0, 'end')
        if not self._headings:
            self._toc_listbox.insert('end', "(无标题)")
            return
        for item in self._headings:
            try:
                if len(item) == 4:
                    level, title, _, _ = item
                elif len(item) == 3:
                    level, title, _ = item
                else:
                    continue
                indent = "  " * (level - 1)
                self._toc_listbox.insert('end', f"{indent}{title}")
            except Exception:
                continue

    def _on_toc_select(self, event):
        """点击大纲项跳转"""
        selection = self._toc_listbox.curselection()
        if not selection or not self._headings:
            return
        idx = selection[0]
        if idx < len(self._headings):
            _, _, source_line = self._headings[idx]
            if self.on_jump_to_line:
                self.on_jump_to_line(source_line)

    def highlight_search_term(self, term: str, case_sensitive: bool = False):
        """在预览区高亮搜索词"""
        self.text.tag_remove("search_highlight", "1.0", "end")
        if not term:
            return

        # 配置高亮样式
        self.text.tag_configure("search_highlight", background="#FFFF00", foreground="#000000")
        
        start_pos = "1.0"
        count = tk.IntVar()
        while True:
            start_pos = self.text.search(
                term, start_pos, stopindex="end", 
                nocase=not case_sensitive, count=count
            )
            if not start_pos:
                break
            end_pos = f"{start_pos}+{count.get()}c"
            self.text.tag_add("search_highlight", start_pos, end_pos)
            start_pos = end_pos

    def _jump_to_anchor(self, anchor_id: str):
        """平滑跳转到内部锚点"""
        if hasattr(self, '_anchor_map') and anchor_id in self._anchor_map:
            target_idx = self._anchor_map[anchor_id]
            # 获取目标行的 y 比例
            line_num = int(target_idx.split('.')[0])
            total_lines = int(self.text.index('end-1c').split('.')[0])
            target_pos = (line_num - 1) / total_lines
            self._smooth_scroll_to(target_pos)
            return "break"
        return None

    def _highlight_preview_line(self, line_num: int):
        """短暂高亮预览区的行"""
        try:
            self.text.tag_configure('jump_highlight', background='#fef3c7')
            self.text.tag_add('jump_highlight', f"{line_num}.0", f"{line_num}.end")
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
        """设置跳转回调函数"""
        self.on_jump_to_line = callback

    def _jump_to_editor_from_context(self):
        """右键菜单触发跳转"""
        try:
            index = self.text.index(tk.INSERT)
            line_num = int(index.split('.')[0])
            source_line = self._find_source_line_for_preview_line(line_num)
            if source_line and self.on_jump_to_line:
                self.on_jump_to_line(source_line)
                self._highlight_preview_line(line_num)
        except Exception:
            pass

    def _on_scrollbar(self, *args):
        """滚动条事件处理"""
        self.text.yview(*args)
        try:
            first = self.text.yview()[0]
            if (getattr(self, '_sync_scroll_enabled', True) and self.on_scroll and not getattr(self, '_scroll_updating', False)):
                self.on_scroll(float(first))
        except Exception:
            pass

    def _on_mousewheel(self, event):
        """预览区平滑滚动实现"""
        delta = -1 * (event.delta // 120)
        current_pos = self.text.yview()[0]
        increment = 0.02 * delta
        target_pos = max(0.0, min(1.0, current_pos + increment))
        if self._performance_mode == 'high':
            self.text.yview_moveto(target_pos)
            return "break"
        self._smooth_scroll_to(target_pos)
        return "break"

    def _smooth_scroll_to(self, target_pos):
        """执行预览区平滑滚动动画（减小抖动）"""
        if self._performance_mode == 'high':
            self.text.yview_moveto(target_pos)
            return
        if self._smooth_scroll_timer:
            self.after_cancel(self._smooth_scroll_timer)
        current_pos = self.text.yview()[0]
        diff = target_pos - current_pos
        if abs(diff) < 0.0008:
            self.text.yview_moveto(target_pos)
            return
        step = diff * 0.12
        # 限制最小步长，避免过冲
        if diff > 0:
            step = max(step, 0.002)
        else:
            step = min(step, -0.002)
        new_pos = current_pos + step
        self.text.yview_moveto(new_pos)
        if self.on_scroll:
            self.on_scroll(new_pos)
        self._smooth_scroll_timer = self.after(10, lambda: self._smooth_scroll_to(target_pos))

    def _on_text_scroll(self, first, last):
        """文本滚动事件处理，同步到编辑器并更新大纲追踪和进度条"""
        self.scrollbar.set(first, last)
        
        # 1. 触发滚动同步回调
        if (getattr(self, '_sync_scroll_enabled', True) and self.on_scroll and not getattr(self, '_scroll_updating', False)):
            try:
                self.on_scroll(float(first))
            except Exception:
                pass
        
        # 2. 更新大纲追踪高亮
        if self._toc_visible:
            delay = 160 if self._performance_mode == 'high' else 50
            self.after(delay, self._track_current_heading)
            
        # 3. 更新阅读进度条
        self._update_reading_progress(float(first))

    def _track_current_heading(self):
        """追踪当前可见区域所在的章节并高亮大纲项"""
        if not self._toc_listbox or not self._headings:
            return
            
        try:
            # 获取可视区域顶部的行号
            first_visible_idx = self.text.index("@0,0")
            visible_line = int(first_visible_idx.split('.')[0])
            
            # 找到当前所在的源码行号
            source_line = self._find_source_line_for_preview_line(visible_line)
            
            # 在 headings 中寻找当前最匹配的标题
            active_idx = -1
            for i, (level, title, line_num) in enumerate(self._headings):
                if line_num <= source_line:
                    active_idx = i
                else:
                    break
            
            # 更新列表框选中状态
            self._toc_listbox.selection_clear(0, 'end')
            if active_idx != -1:
                self._toc_listbox.selection_set(active_idx)
                self._toc_listbox.see(active_idx)
        except Exception:
            pass

    def sync_scroll_to(self, position: float):
        """同步滚动到指定位置"""
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
        """设置是否启用同步滚动"""
        self._sync_scroll_enabled = enabled

    def apply_theme(self, theme_config: dict):
        """应用预览主题"""
        try:
            bg = theme_config.get('background', '#FFFFFF')
            fg = theme_config.get('text_color', '#111827')
            self.text.configure(bg=bg, fg=fg)
            font_family = theme_config.get('font_family', '宋体')
            font_size = theme_config.get('font_size', 16)
            self.text.configure(font=(font_family, font_size))
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
            link_color = theme_config.get('link_color', '#0000FF')
            self.text.tag_configure('link', foreground=link_color)
            code_bg = theme_config.get('code_bg', '#F5F5F5')
            code_color = theme_config.get('code_color', '#1F2937')
            code_font = theme_config.get('code_font', 'Consolas')
            self.text.tag_configure('code', background=code_bg, foreground=code_color, font=(code_font, 10))
            code_block_bg = theme_config.get('code_block_bg', '#FAFAFA')
            code_block_color = theme_config.get('code_block_color', '#1F2937')
            self.text.tag_configure('code_block', background=code_block_bg, foreground=code_block_color, font=(code_font, 10))
            blockquote_bg = theme_config.get('blockquote_bg', '#f9fafb')
            blockquote_color = theme_config.get('blockquote_color', '#6B7280')
            self.text.tag_configure('quote', background=blockquote_bg, foreground=blockquote_color)
        except Exception as e:
            print(f"应用预览主题失败: {e}")
