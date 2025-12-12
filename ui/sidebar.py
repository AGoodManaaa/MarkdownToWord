# -*- coding: utf-8 -*-

import os
import re
import customtkinter as ctk

from ui.theme import COLORS


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
