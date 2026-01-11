# -*- coding: utf-8 -*-

import re
import tkinter as tk
import customtkinter as ctk

from ui.theme import COLORS


class LineNumberedText(ctk.CTkFrame):
    """带行号的文本编辑器 - 精确对齐版"""
    def __init__(self, master, font_size=16, on_scroll=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        
        self.font_size = font_size
        self.on_scroll_callback = on_scroll  # 滚动回调
        
        # 使用原生 tk.Text 而不是 CTkTextbox，以便精确控制
        # 容器框架
        self.container = tk.Frame(self, bg=COLORS['bg_light'])
        self.container.pack(fill='both', expand=True)
        
        # 面包屑导航栏
        self.breadcrumb_frame = tk.Frame(self.container, bg=COLORS['bg_light'], height=24)
        self.breadcrumb_frame.pack(side='top', fill='x', padx=8)
        self.breadcrumb_frame.pack_propagate(False)
        
        self.breadcrumb_label = tk.Label(
            self.breadcrumb_frame,
            text="",
            font=('Microsoft YaHei', 10),
            bg=COLORS['bg_light'],
            fg=COLORS['text_secondary'],
            anchor='w'
        )
        self.breadcrumb_label.pack(side='left', fill='both', expand=True)
        
        # 分隔线
        self.breadcrumb_sep = tk.Frame(self.container, height=1, bg=COLORS['border'])
        self.breadcrumb_sep.pack(side='top', fill='x')
        
        # 内容主区域（行号+装饰线+文本）
        self.main_content_frame = tk.Frame(self.container, bg=COLORS['bg_light'])
        self.main_content_frame.pack(side='top', fill='both', expand=True)
        
        # 行号栏
        self.line_numbers = tk.Text(
            self.main_content_frame,
            width=4,
            padx=4,
            pady=5,
            takefocus=0,
            border=0,
            background=COLORS['line_number_bg'],
            foreground=COLORS['line_number'],
            state='disabled',
            wrap='none',
            font=('Microsoft YaHei', font_size),
            cursor='arrow',
        )
        self.line_numbers.pack(side='left', fill='y')

        # 装饰线画布 (Gutter Decoration)
        self.gutter_canvas = tk.Canvas(
            self.main_content_frame,
            width=4,
            bg=COLORS['bg_light'],
            highlightthickness=0,
            bd=0
        )
        self.gutter_canvas.pack(side='left', fill='y')
        
        # 主文本区 - 使用原生 tk.Text
        self.text_frame = tk.Frame(self.main_content_frame, bg=COLORS['bg_light'])
        self.text_frame.pack(side='left', fill='both', expand=True)
        
        self._textbox = tk.Text(
            self.text_frame,
            font=('Microsoft YaHei', font_size),
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

        self._undo_sep_timer = None

        # 兼容旧属性名
        self.text = self._textbox
        
        # 滚动条 - 使用 CTkScrollbar 保持风格一致
        self.scrollbar = ctk.CTkScrollbar(self.text_frame, command=self._on_scrollbar)
        self.scrollbar.pack(side='right', fill='y')
        
        # 搜索热力图画布 - 紧贴滚动条左侧
        self.heatmap_canvas = tk.Canvas(
            self.text_frame, 
            width=12, 
            bg=COLORS['bg_light'], 
            highlightthickness=0,
            borderwidth=0,
            cursor='hand2'
        )
        self.heatmap_canvas.pack(side='right', fill='y')
        self.heatmap_canvas.bind('<Button-1>', self._on_heatmap_click)
        
        self._textbox.pack(side='left', fill='both', expand=True)
        self._textbox.config(yscrollcommand=self._on_text_scroll)
        
        # 配置当前行高亮标签
        self._textbox.tag_configure("current_line", background=COLORS.get('highlight', '#f0f0f0'))
        # 缩进参考线（使用前导空格着色实现）
        self._textbox.tag_configure("indent_guide", foreground=COLORS.get('border', '#d1d5db'))
        
        # 绑定事件
        self._textbox.bind('<KeyRelease>', self._on_change)
        self._textbox.bind('<ButtonRelease-1>', self._on_change) # 点击时也更新
        self._textbox.bind('<MouseWheel>', self._on_mousewheel)
        self._textbox.bind('<Configure>', self._on_change)
        self.line_numbers.bind('<MouseWheel>', self._on_mousewheel)
        
        self._textbox.bind('<Control-v>', self._on_paste)
        self._textbox.bind('<<Paste>>', self._on_paste)
        
        # 初始化行号
        self.after(50, self._update_line_numbers)
        
        self._typewriter_mode = False  # 打字机模式开关
        self._target_scroll_pos = 0.0
        self._smooth_scroll_timer = None
        
        # 悬浮格式工具栏
        self._selection_toolbar = None
        self._init_selection_toolbar()
        
        # 搜索悬浮层 (Search Overlay)
        self._search_overlay = None
        self._init_search_overlay()
        
        # 图片预览悬浮窗
        self._image_preview_popup = None
        self._textbox.bind('<Motion>', self._on_editor_motion)
    
    def _on_editor_motion(self, event):
        """鼠标移动时检查是否在图片链接上，显示预览"""
        try:
            # 获取鼠标下的位置
            index = self._textbox.index(f"@{event.x},{event.y}")
            line_content = self._textbox.get(f"{index} linestart", f"{index} lineend")
            
            # 查找图片正则
            img_match = re.search(r'!\[.*?\]\((.*?)\)', line_content)
            if img_match:
                url = img_match.group(1)
                # 检查鼠标是否在匹配范围内
                match_start = img_match.start()
                match_end = img_match.end()
                col = int(index.split('.')[1])
                
                if match_start <= col <= match_end:
                    self._show_image_preview(url, event.x_root, event.y_root)
                    return
            
            self._hide_image_preview()
        except:
            self._hide_image_preview()

    def _show_image_preview(self, url, x, y):
        """显示图片缩略图预览"""
        if self._image_preview_popup:
            # 如果已经显示且 URL 相同，只更新位置
            if getattr(self, '_last_preview_url', None) == url:
                self._image_preview_popup.wm_geometry(f"+{x+15}+{y+15}")
                return
            self._hide_image_preview()

        try:
            # 解析路径
            if not os.path.isabs(url) and hasattr(self.winfo_toplevel(), 'current_file'):
                app = self.winfo_toplevel()
                if app.current_file:
                    url = os.path.join(os.path.dirname(app.current_file), url)
            
            if not os.path.exists(url): return

            # 加载缩略图
            from PIL import Image, ImageTk
            img = Image.open(url)
            img.thumbnail((200, 200))
            photo = ImageTk.PhotoImage(img)

            self._image_preview_popup = tk.Toplevel(self)
            self._image_preview_popup.wm_overrideredirect(True)
            self._image_preview_popup.wm_geometry(f"+{x+15}+{y+15}")
            
            lbl = tk.Label(self._image_preview_popup, image=photo, bg='white', borderwidth=1, relief='solid')
            lbl.image = photo # 保持引用
            lbl.pack()
            
            self._last_preview_url = url
        except:
            pass

    def _hide_image_preview(self):
        """隐藏图片预览"""
        if self._image_preview_popup:
            self._image_preview_popup.destroy()
            self._image_preview_popup = None
            self._last_preview_url = None

    def _init_search_overlay(self):
        """初始化编辑器内的搜索悬浮层 (支持替换)"""
        self._search_overlay = tk.Frame(
            self.text_frame,
            bg=COLORS.get('bg_card', '#ffffff'),
            highlightthickness=1,
            highlightbackground=COLORS.get('border', '#e5e7eb'),
            padx=10, pady=8
        )
        
        # 1. 搜索行
        search_row = tk.Frame(self._search_overlay, bg=COLORS.get('bg_card', '#ffffff'))
        search_row.pack(fill='x', pady=(0, 5))
        
        self._search_input = tk.Entry(
            search_row, 
            width=20, 
            font=('Microsoft YaHei', 10),
            bd=0,
            highlightthickness=1,
            highlightcolor=COLORS.get('primary', '#3b82f6'),
            highlightbackground=COLORS.get('border', '#e5e7eb')
        )
        self._search_input.pack(side='left', padx=(0, 5))
        self._search_input.bind('<KeyRelease>', self._on_search_overlay_change)
        self._search_input.bind('<Return>', lambda e: self._search_overlay_nav(1))
        self._search_input.bind('<Escape>', lambda e: self.hide_search_overlay())
        
        # 导航按钮
        tk.Button(search_row, text="↑", command=lambda: self._search_overlay_nav(-1), font=('Arial', 10), bg=COLORS.get('bg_card'), bd=0, cursor='hand2').pack(side='left', padx=2)
        tk.Button(search_row, text="↓", command=lambda: self._search_overlay_nav(1), font=('Arial', 10), bg=COLORS.get('bg_card'), bd=0, cursor='hand2').pack(side='left', padx=2)
        tk.Button(search_row, text="✕", command=self.hide_search_overlay, font=('Arial', 10), bg=COLORS.get('bg_card'), bd=0, cursor='hand2').pack(side='right', padx=(5, 0))
        
        # 2. 替换行
        self._replace_row = tk.Frame(self._search_overlay, bg=COLORS.get('bg_card', '#ffffff'))
        self._replace_row.pack(fill='x', pady=(5, 0))
        
        self._replace_input = tk.Entry(
            self._replace_row, 
            width=20, 
            font=('Microsoft YaHei', 10),
            bd=0,
            highlightthickness=1,
            highlightcolor=COLORS.get('success', '#10b981'),
            highlightbackground=COLORS.get('border', '#e5e7eb')
        )
        self._replace_input.pack(side='left', padx=(0, 5))
        
        tk.Button(self._replace_row, text="替换", command=self._search_overlay_replace, font=('Microsoft YaHei', 8), bg=COLORS.get('bg_card'), bd=1, relief='flat', cursor='hand2').pack(side='left', padx=2)
        tk.Button(self._replace_row, text="全部", command=self._search_overlay_replace_all, font=('Microsoft YaHei', 8), bg=COLORS.get('bg_card'), bd=1, relief='flat', cursor='hand2').pack(side='left', padx=2)

        # 3. 结果统计
        self._search_stats_lbl = tk.Label(self._search_overlay, text="", font=('Microsoft YaHei', 8), bg=COLORS.get('bg_card'), fg=COLORS.get('text_secondary'))
        self._search_stats_lbl.pack(fill='x', pady=(4, 0))

    def _search_overlay_replace(self):
        """执行当前匹配项的替换"""
        if not hasattr(self, '_search_matches') or not self._search_matches:
            return
        
        try:
            start, end = self._search_matches[self._current_match_idx]
            new_text = self._replace_input.get()
            self._textbox.delete(start, end)
            self._textbox.insert(start, new_text)
            self._on_search_overlay_change() # 重新扫描
        except:
            pass

    def _search_overlay_replace_all(self):
        """执行全部替换"""
        term = self._search_input.get()
        if not term: return
        
        try:
            new_text = self._replace_input.get()
            content = self._textbox.get("1.0", "end-1c")
            # 使用正则全局替换
            replaced_content = re.sub(re.escape(term), new_text, content, flags=re.IGNORECASE)
            
            if content != replaced_content:
                self._textbox.delete("1.0", "end")
                self._textbox.insert("1.0", replaced_content)
                self._on_search_overlay_change()
                if hasattr(self.winfo_toplevel(), 'update_status'):
                    self.winfo_toplevel().update_status(f"✅ 已全部替换为: {new_text}", is_temp=True)
        except Exception as e:
            print(f"Replace all error: {e}")

    def show_search_overlay(self, initial_text=""):
        """显示搜索悬浮层"""
        self._search_overlay.place(relx=1.0, rely=0, anchor='ne', x=-30, y=10)
        self._search_input.focus_set()
        if initial_text:
            self._search_input.delete(0, 'end')
            self._search_input.insert(0, initial_text)
            self._on_search_overlay_change()

    def hide_search_overlay(self):
        """隐藏搜索悬浮层"""
        self._search_overlay.place_forget()
        self._textbox.tag_remove("search_highlight", "1.0", "end")
        self._textbox.tag_remove("current_match", "1.0", "end")
        self.update_search_heatmap([])

    def _on_search_overlay_change(self, event=None):
        """处理搜索层内容变化"""
        term = self._search_input.get()
        self._textbox.tag_remove("search_highlight", "1.0", "end")
        if not term:
            self._search_stats_lbl.configure(text="")
            self.update_search_heatmap([])
            return
            
        # 使用正则表达式查找
        content = self._textbox.get("1.0", "end-1c")
        matches = []
        try:
            for m in re.finditer(re.escape(term), content, re.IGNORECASE):
                start = self._index_to_pos(m.start(), content)
                end = self._index_to_pos(m.end(), content)
                matches.append((start, end))
                self._textbox.tag_add("search_highlight", start, end)
            
            self._textbox.tag_configure("search_highlight", background="#FFFF00", foreground="#000000")
            self._textbox.tag_configure("current_match", background="#FF8C00", foreground="#FFFFFF")
            
            self._search_matches = matches
            self._current_match_idx = 0
            if matches:
                self._search_overlay_nav(0)
                self.update_search_heatmap(matches)
            else:
                self._search_stats_lbl.configure(text="未找到匹配")
                self.update_search_heatmap([])
        except:
            pass

    def _search_overlay_nav(self, delta):
        """在匹配结果中导航"""
        if not hasattr(self, '_search_matches') or not self._search_matches:
            return
            
        self._textbox.tag_remove("current_match", "1.0", "end")
        self._current_match_idx = (self._current_match_idx + delta) % len(self._search_matches)
        
        start, end = self._search_matches[self._current_match_idx]
        self._textbox.tag_add("current_match", start, end)
        self._textbox.see(start)
        self._search_stats_lbl.configure(text=f"{self._current_match_idx+1} / {len(self._search_matches)}")

    def _init_selection_toolbar(self):
        """初始化选中文本后的悬浮工具栏 (增加统计显示)"""
        self._selection_toolbar = tk.Frame(
            self.text_frame, 
            bg=COLORS.get('bg_card', '#ffffff'),
            highlightthickness=1,
            highlightbackground=COLORS.get('border', '#e5e7eb'),
            padx=2, pady=2
        )
        
        # 按钮容器
        btn_frame = tk.Frame(self._selection_toolbar, bg=COLORS.get('bg_card', '#ffffff'))
        btn_frame.pack(side='top', fill='x')
        
        # 定义操作按钮
        btns = [
            ("B", self._format_bold, "加粗"),
            ("I", self._format_italic, "斜体"),
            ("`", self._format_code, "行内代码"),
            ("🔗", self._format_link, "链接")
        ]
        
        for text, cmd, tip in btns:
            btn = tk.Button(
                btn_frame,
                text=text,
                command=cmd,
                font=('Microsoft YaHei', 9, 'bold'),
                bg=COLORS.get('bg_card', '#ffffff'),
                fg=COLORS.get('text_primary', '#111827'),
                activebackground=COLORS.get('highlight', '#f3f4f6'),
                borderwidth=0,
                padx=6, pady=2,
                cursor='hand2'
            )
            btn.pack(side='left', padx=1)
            
        # 统计显示层
        self._selection_stats_lbl = tk.Label(
            self._selection_toolbar,
            text="",
            font=('Microsoft YaHei', 8),
            bg=COLORS.get('bg_card', '#ffffff'),
            fg=COLORS.get('text_secondary', '#6b7280'),
            pady=2
        )
        self._selection_stats_lbl.pack(side='bottom', fill='x')

    def _format_bold(self): self._wrap_selection("**", "**")
    def _format_italic(self): self._wrap_selection("*", "*")
    def _format_code(self): self._wrap_selection("`", "`")
    def _format_link(self): self._wrap_selection("[", "](url)")

    def _wrap_selection(self, prefix, suffix):
        """用指定字符包裹选中文本"""
        try:
            sel_start = self._textbox.index(tk.SEL_FIRST)
            sel_end = self._textbox.index(tk.SEL_LAST)
            content = self._textbox.get(sel_start, sel_end)
            self._textbox.delete(sel_start, sel_end)
            self._textbox.insert(sel_start, f"{prefix}{content}{suffix}")
            self._selection_toolbar.place_forget()
            self._on_change()
        except tk.TclError:
            pass

    def _update_selection_toolbar(self):
        """检查选中状态并更新工具栏位置，增加避让逻辑和统计"""
        try:
            # 检查是否有选中
            try:
                sel_start = self._textbox.index(tk.SEL_FIRST)
                sel_end = self._textbox.index(tk.SEL_LAST)
            except tk.TclError:
                self._selection_toolbar.place_forget()
                return
            
            if sel_start == sel_end:
                self._selection_toolbar.place_forget()
                return

            selected_text = self._textbox.get(sel_start, sel_end)
            char_count = len(selected_text)
            word_count = len(re.findall(r'\w+', selected_text))
            self._selection_stats_lbl.configure(text=f"{char_count} 字 | {word_count} 词")

            # 获取选中文本的坐标位置
            bbox = self._textbox.bbox(sel_start)
            if bbox:
                x, y, w, h = bbox
                
                # 工具栏尺寸（预估）
                tw = 160
                th = 50 # 增加高度以容纳统计
                
                # 默认显示在上方
                tx = x
                ty = y - th - 5
                
                # 边界检查：如果上方空间不足，显示在下方
                if ty < 0:
                    ty = y + h + 5
                
                # 横向边界检查：防止超出右侧
                max_x = self.text_frame.winfo_width() - tw
                if tx > max_x:
                    tx = max_x
                tx = max(0, tx)
                
                self._selection_toolbar.place(x=tx, y=ty)
            else:
                self._selection_toolbar.place_forget()
        except Exception:
            self._selection_toolbar.place_forget()

    def _on_paste(self, event=None):
        """智能粘贴功能：处理链接、表格、HTML 和剪贴板图片"""
        try:
            # 1. 优先尝试从剪贴板读取图片
            from PIL import ImageGrab
            try:
                img = ImageGrab.grabclipboard()
                if img:
                    self._handle_clipboard_image(img)
                    return "break"
            except Exception:
                pass

            # 2. 获取文本剪贴板内容
            clipboard = self.clipboard_get()
            if not clipboard: return
            
            # 3. 检查是否是 URL
            url_pattern = r'^https?://[^\s/$.?#].[^\s]*$'
            if re.match(url_pattern, clipboard.strip()):
                try:
                    sel_start = self._textbox.index(tk.SEL_FIRST)
                    sel_end = self._textbox.index(tk.SEL_LAST)
                    selected_text = self._textbox.get(sel_start, sel_end)
                    if selected_text:
                        self._textbox.delete(sel_start, sel_end)
                        self._textbox.insert(sel_start, f"[{selected_text}]({clipboard.strip()})")
                        self._on_change()
                        return "break"
                except tk.TclError:
                    pass
            
            # 2. 检查是否是表格数据
            if '\t' in clipboard and '\n' in clipboard:
                lines = [line.strip() for line in clipboard.strip().split('\n') if line.strip()]
                if len(lines) > 1:
                    table_rows = [line.split('\t') for line in lines]
                    col_counts = [len(row) for row in table_rows]
                    if max(col_counts) > 1 and len(set(col_counts)) <= 2:
                        md_table = self._convert_to_md_table(table_rows)
                        self._textbox.insert(tk.INSERT, md_table)
                        self._on_change()
                        return "break"

            # 3. 检查是否包含 HTML 标签并尝试转换
            if '<' in clipboard and '>' in clipboard and ('</' in clipboard or '/>' in clipboard):
                # 如果检测到明显的 HTML 结构，尝试做基础转换
                try:
                    md_text = self._basic_html_to_md(clipboard)
                    if md_text != clipboard: # 如果转换成功（内容有变化）
                        self._textbox.insert(tk.INSERT, md_text)
                        self._on_change()
                        if hasattr(self.winfo_toplevel(), 'update_status'):
                            self.winfo_toplevel().update_status("🪄 已自动将粘贴的 HTML 转换为 Markdown", is_temp=True)
                        return "break"
                except Exception:
                    pass
                        
        except Exception:
            pass
        return None

    def _basic_html_to_md(self, html_content):
        """基础 HTML 转 Markdown (非正则版，使用内建逻辑)"""
        import re
        text = html_content
        # 1. 替换标题
        for i in range(6, 0, -1):
            text = re.sub(f'<h{i}[^>]*>(.*?)</h{i}>', lambda m: f"\n{'#'*i} {m.group(1)}\n", text, flags=re.S|re.I)
        # 2. 替换格式标签
        text = re.sub(r'<(b|strong)[^>]*>(.*?)</\1>', r'**\2**', text, flags=re.S|re.I)
        text = re.sub(r'<(i|em)[^>]*>(.*?)</\1>', r'*\2*', text, flags=re.S|re.I)
        text = re.sub(r'<code[^>]*>(.*?)</code>', r'`\1`', text, flags=re.S|re.I)
        # 3. 链接和图片
        text = re.sub(r'<a[^>]*href=["\'](.*?)["\'][^>]*>(.*?)</a>', r'[\2](\1)', text, flags=re.S|re.I)
        text = re.sub(r'<img[^>]*src=["\'](.*?)["\'][^>]*alt=["\'](.*?)["\'][^>]*>', r'![\2](\1)', text, flags=re.S|re.I)
        text = re.sub(r'<img[^>]*alt=["\'](.*?)["\'][^>]*src=["\'](.*?)["\'][^>]*>', r'![\1](\2)', text, flags=re.S|re.I)
        # 4. 列表
        text = re.sub(r'<li[^>]*>(.*?)</li>', r'\n- \1', text, flags=re.S|re.I)
        # 5. 清理剩余标签
        text = re.sub(r'<br\s*/?>', r'\n', text, flags=re.I)
        text = re.sub(r'<(p|div|ul|ol)[^>]*>', r'\n', text, flags=re.I)
        text = re.sub(r'</(p|div|ul|ol)>', r'\n', text, flags=re.I)
        # 6. 去掉所有其他标签
        text = re.sub(r'<[^>]+>', '', text)
        # 7. 处理实体
        text = text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        return text.strip()

    def _handle_clipboard_image(self, img):
        """处理剪贴板中的图片并存入 assets"""
        try:
            # 检查是否有打开的文件以确定路径
            app = self.winfo_toplevel()
            if not hasattr(app, 'current_file') or not app.current_file:
                tk.messagebox.showinfo('提示', '请先保存 Markdown 文件，以便确定图片 assets 目录的位置。')
                return

            # 确定 assets 目录
            doc_dir = os.path.dirname(app.current_file)
            assets_dir = os.path.join(doc_dir, 'assets')
            if not os.path.exists(assets_dir):
                os.makedirs(assets_dir)

            # 生成文件名
            import time
            timestamp = int(time.time() * 1000)
            filename = f"paste_image_{timestamp}.png"
            dest_path = os.path.join(assets_dir, filename)

            # 保存图片
            img.save(dest_path, "PNG")

            # 插入链接
            rel_path = f"assets/{filename}"
            md_code = f"![剪贴板图片]({rel_path})\n"
            self._textbox.insert(tk.INSERT, md_code)
            self._on_change()
            
            if hasattr(app, 'update_status'):
                app.update_status(f"✅ 已保存并插入剪贴板图片", is_temp=True, pulse=True)
        except Exception as e:
            print(f"Paste image error: {e}")

    def _convert_to_md_table(self, rows):
        """将嵌套列表转换为 Markdown 表格字符串"""
        if not rows: return ""
        
        # 确定最大列数
        num_cols = max(len(row) for row in rows)
        
        md_lines = []
        # 表头
        headers = rows[0]
        md_lines.append("| " + " | ".join(headers + [""] * (num_cols - len(headers))) + " |")
        # 分隔线
        md_lines.append("| " + " | ".join(["---"] * num_cols) + " |")
        # 数据行
        for row in rows[1:]:
            md_lines.append("| " + " | ".join(row + [""] * (num_cols - len(row))) + " |")
            
        return "\n" + "\n".join(md_lines) + "\n"

    def set_typewriter_mode(self, enabled: bool):
        """设置打字机模式（仅光标居中，不做淡化）"""
        self._typewriter_mode = enabled
        try:
            # 立即刷新当前行高亮
            self._highlight_current_line()
            # 打字机模式：仅保持光标居中，移除淡化逻辑
            if enabled:
                self._textbox.tag_remove("zen_dim", "1.0", "end")
                self._center_cursor_v2()
            else:
                self._textbox.tag_remove("zen_dim", "1.0", "end")
        except Exception:
            pass

    def _center_cursor(self):
        """将光标所在行滚动到屏幕中央"""
        if not self._typewriter_mode:
            return
        try:
            # 延迟执行确保位置准确
            self._textbox.see("insert")
            # 使用 yview_scroll 或 yview_moveto 来居中，这里使用 yview('insert') 配合偏移
            # 获取当前可见范围
            # tk.Text.yview() 返回 (top, bottom) 百分比
            # 这里简单实现：如果光标行不是中央，则调整 yview
            # 更靠谱的方案是获取总行数和当前行，计算 moveto
            pass
        except Exception:
            pass

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
        """鼠标滚轮平滑滚动实现"""
        # 计算滚动的单位增量 (通常每次滚动 120 单位)
        delta = -1 * (event.delta // 120)
        
        # 获取当前位置
        current_pos = self._textbox.yview()[0]
        # 计算目标位置（近似增量，这里设为每次滚动约 3 行的高度比例）
        # 实际滚动量取决于内容长度，这里采用简单的增量方式
        increment = 0.02 * delta # 假设一次滚动占总长的 2%
        
        target_pos = max(0.0, min(1.0, current_pos + increment))
        self._smooth_scroll_to(target_pos)
        return "break"

    def _smooth_scroll_to(self, target_pos):
        """执行平滑滚动动画"""
        if self._smooth_scroll_timer:
            self.after_cancel(self._smooth_scroll_timer)
            
        current_pos = self._textbox.yview()[0]
        diff = target_pos - current_pos
        
        if abs(diff) < 0.001:
            self._textbox.yview_moveto(target_pos)
            self.line_numbers.yview_moveto(target_pos)
            return

        # 每次移动 20% 的剩余距离
        step = diff * 0.2
        new_pos = current_pos + step
        
        self._textbox.yview_moveto(new_pos)
        self.line_numbers.yview_moveto(new_pos)
        
        # 通知外部滚动回调
        if self.on_scroll_callback:
            self.on_scroll_callback(new_pos)
            
        self._smooth_scroll_timer = self.after(10, lambda: self._smooth_scroll_to(target_pos))
    
    def _on_change(self, event=None):
        """内容变化时更新行号、高亮、参考线、面包屑等"""
        self.after(5, self._update_line_numbers)
        self.after(5, self._highlight_current_line)
        self.after(5, self._update_indent_guides)
        self.after(5, self._update_breadcrumbs)
        self.after(50, self._update_selection_toolbar)
        
        # 触发智能编辑器的更新（如括号染色）
        if hasattr(self, 'smart_editor_ref') and self.smart_editor_ref:
            self.after(10, self.smart_editor_ref.bracket_colorizer.update)
        
        # 打字机模式：保持光标居中
        if self._typewriter_mode:
            self.after(10, self._center_cursor_v2)

        # 让撤销更“一级一级”：对连续输入做轻量防抖后插入 undo 分隔点
        try:
            if self._undo_sep_timer is not None:
                self.after_cancel(self._undo_sep_timer)
        except Exception:
            pass
        try:
            self._undo_sep_timer = self.after(180, self._insert_undo_separator)
        except Exception:
            self._undo_sep_timer = None

    def _highlight_current_line(self):
        """高亮当前行（移除打字机淡化效果）"""
        try:
            self._textbox.tag_remove("current_line", "1.0", "end")
            self._textbox.tag_add("current_line", "insert linestart", "insert lineend+1c")
            # 确保不存在残留的淡化标签
            self._textbox.tag_remove("zen_dim", "1.0", "end")
        except Exception:
            pass

    def _update_breadcrumbs(self):
        """更新面包屑导航，并支持点击跳转"""
        try:
            cursor_idx = self._textbox.index("insert")
            current_line = int(cursor_idx.split('.')[0])
            
            content = self._textbox.get("1.0", f"{current_line}.end").split("\n")
            
            # 清除旧的按钮
            for child in self.breadcrumb_frame.winfo_children():
                child.destroy()
            
            # 向上搜索各级标题
            current_levels = [] # [(level, title, line_num)]
            
            latest_levels = {1: None, 2: None, 3: None, 4: None, 5: None, 6: None}
            
            for i in range(current_line):
                line = content[i].strip()
                if line.startswith("#"):
                    match = re.match(r'^(#+)\s+(.*)', line)
                    if match:
                        level = len(match.group(1))
                        title = match.group(2)
                        latest_levels[level] = (title, i + 1)
                        # 清除更深级别的标题
                        for l in range(level + 1, 7):
                            latest_levels[l] = None
            
            # 构建按钮序列
            hierarchy = []
            for l in range(1, 7):
                if latest_levels[l]:
                    hierarchy.append(latest_levels[l])
            
            if not hierarchy:
                lbl = tk.Label(self.breadcrumb_frame, text="正文", font=('Microsoft YaHei', 9),
                             bg=COLORS['bg_light'], fg=COLORS['text_secondary'])
                lbl.pack(side='left')
                return

            for i, (title, line_num) in enumerate(hierarchy):
                if i > 0:
                    sep = tk.Label(self.breadcrumb_frame, text=" > ", font=('Microsoft YaHei', 9),
                                 bg=COLORS['bg_light'], fg=COLORS['text_secondary'])
                    sep.pack(side='left')
                
                btn = tk.Label(
                    self.breadcrumb_frame,
                    text=title[:20] + "..." if len(title) > 20 else title,
                    font=('Microsoft YaHei', 9),
                    bg=COLORS['bg_light'],
                    fg=COLORS['text_secondary'],
                    cursor='hand2'
                )
                btn.pack(side='left')
                btn.bind('<Enter>', lambda e, b=btn: b.configure(fg=COLORS.get('primary', '#3b82f6')))
                btn.bind('<Leave>', lambda e, b=btn: b.configure(fg=COLORS['text_secondary']))
                btn.bind('<Button-1>', lambda e, ln=line_num: self._jump_to_line(ln))
                # 增加右键菜单
                btn.bind('<Button-3>', lambda e, ln=line_num, t=title: self._show_breadcrumb_menu(e, ln, t))
                
        except Exception:
            pass

    def _jump_to_line(self, line_num):
        """跳转到指定行"""
        try:
            self._textbox.see(f"{line_num}.0")
            self._textbox.mark_set("insert", f"{line_num}.0")
            self._textbox.focus_set()
            self._on_change()
        except Exception:
            pass

    def _show_breadcrumb_menu(self, event, line_num, title):
        """显示面包屑导航右键菜单"""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"跳转到: {title[:15]}...", command=lambda: self._jump_to_line(line_num))
        menu.add_separator()
        menu.add_command(label="查看章节统计", command=lambda: self._show_section_stats(line_num, title))
        menu.add_command(label="复制此章节内容", command=lambda: self._copy_section(line_num))
        
        # 增加导出章节功能
        export_menu = tk.Menu(menu, tearoff=0)
        export_menu.add_command(label="导出为 Word", command=lambda: self._export_section(line_num, 'word'))
        export_menu.add_command(label="导出为 PDF", command=lambda: self._export_section(line_num, 'pdf'))
        menu.add_cascade(label="导出此章节", menu=export_menu)
        
        # 如果支持折叠功能，添加折叠选项
        if hasattr(self.master.master.master, 'code_folding'):
            menu.add_command(label="折叠此章节", command=lambda: self._fold_section(line_num))
            
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _export_section(self, line_num, fmt):
        """导出当前章节内容"""
        try:
            content = self._textbox.get("1.0", "end-1c").split("\n")
            start_line = line_num - 1
            start_header_level = self._get_heading_level(content[start_line])
            
            if start_header_level == 0: return
            
            end_line = len(content)
            for i in range(start_line + 1, len(content)):
                level = self._get_heading_level(content[i])
                if level > 0 and level <= start_header_level:
                    end_line = i
                    break
            
            section_text = "\n".join(content[start_line:end_line])
            
            app = self.winfo_toplevel()
            if hasattr(app, 'export_to_word') and fmt == 'word':
                # 这里由于 app.export_to_word 是从整个文档导出的，需要临时创建一个导出逻辑或传入内容
                from ui.export_helpers import do_export_for_app
                do_export_for_app(app, section_text, "default", "A4")
            elif hasattr(app, 'pdf_export_feature') and fmt == 'pdf':
                app.pdf_export_feature.export_to_pdf(content=section_text)
        except Exception as e:
            print(f"Export section error: {e}")

    def _show_section_stats(self, line_num, title):
        """显示章节详细统计信息"""
        try:
            content = self._textbox.get("1.0", "end-1c").split("\n")
            start_line = line_num - 1
            start_header_level = self._get_heading_level(content[start_line])
            
            if start_header_level == 0: return
            
            end_line = len(content)
            for i in range(start_line + 1, len(content)):
                level = self._get_heading_level(content[i])
                if level > 0 and level <= start_header_level:
                    end_line = i
                    break
            
            section_text = "\n".join(content[start_line:end_line])
            
            # 使用已有的 StatisticsDetailFeature 计算
            app = self.winfo_toplevel()
            if hasattr(app, 'statistics_detail'):
                stats = app.statistics_detail.calculate_statistics(section_text)
                
                # 弹窗显示统计信息
                msg = f"章节: {title}\n"
                msg += "=" * 20 + "\n"
                msg += f"字数 (不含空格): {stats.chars_no_spaces}\n"
                msg += f"总字符数: {stats.total_chars}\n"
                msg += f"中文字符: {stats.chinese_chars}\n"
                msg += f"英文单词: {stats.english_words}\n"
                msg += f"行数: {end_line - start_line}\n"
                msg += f"预计阅读时间: {stats.reading_time_minutes} 分钟"
                
                tk.messagebox.showinfo("📊 章节统计", msg)
        except Exception as e:
            print(f"Section stats error: {e}")

    def _copy_section(self, line_num):
        """从指定行开始复制整个章节内容"""
        try:
            content = self._textbox.get("1.0", "end-1c").split("\n")
            start_line = line_num - 1
            start_header_level = self._get_heading_level(content[start_line])
            
            if start_header_level == 0: return
            
            end_line = len(content)
            for i in range(start_line + 1, len(content)):
                level = self._get_heading_level(content[i])
                if level > 0 and level <= start_header_level:
                    end_line = i
                    break
            
            section_text = "\n".join(content[start_line:end_line])
            self.clipboard_clear()
            self.clipboard_append(section_text)
            
            # 反馈
            app = self.winfo_toplevel()
            if hasattr(app, 'update_status'):
                app.update_status("📋 章节内容已复制到剪贴板", is_temp=True)
        except Exception:
            pass

    def _get_heading_level(self, line):
        """获取行标题等级"""
        stripped = line.strip()
        if stripped.startswith("#"):
            match = re.match(r'^(#+)', stripped)
            if match:
                return len(match.group(1))
        return 0

    def _fold_section(self, line_num):
        """通过代码折叠功能折叠当前章节"""
        try:
            # 这里的 master 寻找逻辑可能比较脆弱，建议通过 app 引用更稳健
            app = self.winfo_toplevel()
            if hasattr(app, 'code_folding'):
                app.code_folding.toggle_fold(line_num)
        except Exception:
            pass

    def _update_indent_guides(self):
        """为缩进层级着色，形成参考线效果"""
        try:
            self._textbox.tag_remove("indent_guide", "1.0", "end")
            content = self._textbox.get("1.0", "end-1c").split("\n")
            for line_idx, line in enumerate(content, start=1):
                if not line:
                    continue
                leading = len(line) - len(line.lstrip(" "))
                if leading <= 1:
                    continue
                # 每4空格为一级，引导线给前导空格着色
                for pos in range(0, leading, 4):
                    start = f"{line_idx}.{pos}"
                    end = f"{line_idx}.{min(pos + 4, leading)}"
                    self._textbox.tag_add("indent_guide", start, end)
        except Exception:
            pass

    def _center_cursor_v2(self):
        """打字机模式核心实现：平滑地将光标行置于可视区域中央"""
        if not self._typewriter_mode:
            return
        try:
            # 获取光标位置比率
            cursor_idx = self._textbox.index("insert")
            line_num = int(cursor_idx.split('.')[0])
            total_lines = int(self._textbox.index('end-1c').split('.')[0])
            
            # 计算可视行数
            top_idx = self._textbox.index("@0,0")
            bottom_idx = self._textbox.index(f"@0,{self._textbox.winfo_height()}")
            visible_lines = int(bottom_idx.split('.')[0]) - int(top_idx.split('.')[0])
            
            if visible_lines > 5:
                # 目标：让光标行处于可视区域的 40% - 50% 位置
                # 计算目标 top_line
                target_top_line = max(1, line_num - int(visible_lines * 0.45))
                target_pos = (target_top_line - 1) / total_lines
                
                # 使用已有的平滑滚动机制
                self._smooth_scroll_to(target_pos)
        except Exception:
            pass

    def _insert_undo_separator(self):
        self._undo_sep_timer = None
        try:
            self._textbox.edit_separator()
        except Exception:
            pass
    
    def _update_line_numbers(self):
        """优化后的行号显示：仅在行数变化时重绘，或滚动时同步位置，并显示折叠图标与装饰线"""
        try:
            # 获取当前文本行数
            content = self._textbox.get('1.0', 'end-1c')
            lines = content.split('\n')
            line_count = len(lines)
            
            # 只有当行数真正改变时才更新内容，减少 DOM 操作
            if not hasattr(self, '_last_line_count') or self._last_line_count != line_count:
                self.line_numbers.config(state='normal')
                self.line_numbers.delete('1.0', 'end')
                
                # 获取折叠区域信息（如果可用）
                fold_regions = {}
                app = self.winfo_toplevel()
                if hasattr(app, 'code_folding'):
                    fold_regions = {r.start_line: r for r in app.code_folding.get_fold_regions()}

                for i in range(1, line_count + 1):
                    prefix = " "
                    if i in fold_regions:
                        prefix = "▶" if fold_regions[i].is_folded else "▼"
                    
                    # 使用固定宽度的行号显示，左侧预留图标位
                    line_str = f"{prefix}{str(i).rjust(3)}\n"
                    self.line_numbers.insert('end', line_str)
                    
                    # 为图标位增加标签，支持点击
                    if i in fold_regions:
                        tag_name = f"fold_mark_{i}"
                        start_idx = f"{i}.0"
                        end_idx = f"{i}.1"
                        self.line_numbers.tag_add(tag_name, start_idx, end_idx)
                        self.line_numbers.tag_bind(tag_name, '<Button-1>', 
                                                 lambda e, ln=i: app.code_folding.toggle_fold(ln))
                
                self.line_numbers.config(state='disabled')
                self._last_line_count = line_count
            
            # 始终同步滚动位置
            self.line_numbers.yview_moveto(self._textbox.yview()[0])
            # 更新装饰线（如引用块边框）
            self._update_gutter_decorations(lines)
        except Exception:
            pass

    def _update_gutter_decorations(self, lines):
        """在 gutter 中绘制装饰元素，如引用块的侧边线"""
        self.gutter_canvas.delete("all")
        try:
            # 获取可见区域行号
            top_line = int(self._textbox.index("@0,0").split('.')[0])
            bottom_line = int(self._textbox.index(f"@0,{self._textbox.winfo_height()}").split('.')[0])
            
            for i in range(top_line, min(bottom_line + 1, len(lines) + 1)):
                line_content = lines[i-1].strip()
                if line_content.startswith(">"):
                    bbox = self._textbox.bbox(f"{i}.0")
                    if bbox:
                        x, y, w, h = bbox
                        # 在对应位置画一条细线
                        self.gutter_canvas.create_line(
                            2, y, 2, y + h,
                            fill="#cbd5e1", width=2, tags="quote_line"
                        )
        except:
            pass
    
    def _on_heatmap_click(self, event):
        """点击热力图跳转到最近的匹配项"""
        if not hasattr(self, '_heatmap_matches') or not self._heatmap_matches:
            return
            
        try:
            canvas_height = self.heatmap_canvas.winfo_height()
            total_lines = int(self._textbox.index('end-1c').split('.')[0])
            
            # 计算点击位置对应的行号
            clicked_ratio = event.y / canvas_height
            target_line = int(clicked_ratio * total_lines)
            
            # 寻找最近的匹配项
            nearest_match = min(self._heatmap_matches, 
                               key=lambda m: abs(int(m[0].split('.')[0]) - target_line))
            
            self._jump_to_line(nearest_match[0].split('.')[0])
        except Exception:
            pass

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
    
    def bind(self, event, callback, add='+'):
        """绑定事件"""
        self._textbox.bind(event, callback, add=add)
    
    def set_font_size(self, size):
        """设置字体大小"""
        self.font_size = size
        self._textbox.configure(font=('Microsoft YaHei', size))
        self.line_numbers.configure(font=('Microsoft YaHei', size))
        self._update_line_numbers()

    def update_search_heatmap(self, matches: list):
        """更新搜索热力图显示匹配项位置"""
        self.heatmap_canvas.delete("all")
        self._heatmap_matches = matches # 保存匹配项用于点击跳转
        if not matches:
            return
            
        try:
            # 获取文本总长度（字符数或行数）
            total_lines = int(self._textbox.index('end-1c').split('.')[0])
            canvas_height = self.heatmap_canvas.winfo_height()
            
            if canvas_height <= 1:
                self.after(100, lambda: self.update_search_heatmap(matches))
                return

            # 对每一处匹配，在画布上画一个小横条
            for start_idx, end_idx in matches:
                line_num = int(start_idx.split('.')[0])
                # 计算垂直比例位置
                y_pos = (line_num / total_lines) * canvas_height
                
                # 画出标记（橙色）
                self.heatmap_canvas.create_line(
                    2, y_pos, 10, y_pos, 
                    fill="#FF8C00", width=2
                )
        except Exception:
            pass
