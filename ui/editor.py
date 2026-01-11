# -*- coding: utf-8 -*-

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
            font=('Microsoft YaHei', font_size),
            cursor='arrow',
        )
        self.line_numbers.pack(side='left', fill='y')
        
        # 主文本区 - 使用原生 tk.Text
        self.text_frame = tk.Frame(self.container, bg=COLORS['bg_light'])
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
        """内容变化时更新行号和当前行高亮"""
        self.after(5, self._update_line_numbers)
        self.after(5, self._highlight_current_line)
        self.after(5, self._update_indent_guides)

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
        """高亮当前行"""
        try:
            self._textbox.tag_remove("current_line", "1.0", "end")
            self._textbox.tag_add("current_line", "insert linestart", "insert lineend+1c")
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

    def _insert_undo_separator(self):
        self._undo_sep_timer = None
        try:
            self._textbox.edit_separator()
        except Exception:
            pass
    
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
    
    def bind(self, event, callback, add='+'):
        """绑定事件"""
        self._textbox.bind(event, callback, add=add)
    
    def set_font_size(self, size):
        """设置字体大小"""
        self.font_size = size
        self._textbox.configure(font=('Microsoft YaHei', size))
        self.line_numbers.configure(font=('Microsoft YaHei', size))
        self._update_line_numbers()
