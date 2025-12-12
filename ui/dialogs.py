# -*- coding: utf-8 -*-

import tkinter as tk
import customtkinter as ctk

from ui.theme import COLORS


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
