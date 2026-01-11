# -*- coding: utf-8 -*-
"""
全局搜索替换功能
支持正则表达式、高亮匹配、批量替换
"""

import re
import customtkinter as ctk
from tkinter import END
from typing import List, Tuple, Optional
from ui.dialog_utils import set_dialog_icon


class GlobalSearchReplaceFeature:
    """全局搜索替换功能"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.matches: List[Tuple[str, str]] = []  # [(start, end), ...]
        self.current_match_index = 0
        self.highlight_tag = "search_highlight"
        self.current_tag = "current_match"
    
    def show_dialog(self):
        """显示搜索替换对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🔎 搜索和替换")
        self.dialog.geometry("500x400")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 500) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 400) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 搜索输入
        ctk.CTkLabel(main_frame, text="搜索:", anchor="w").pack(fill="x", pady=(0, 5))
        self.search_entry = ctk.CTkEntry(main_frame, width=400, height=35)
        self.search_entry.pack(fill="x", pady=(0, 10))
        self.search_entry.bind("<Return>", lambda e: self.find_next())
        self.search_entry.bind("<KeyRelease>", self._on_search_change)
        
        # 替换输入
        ctk.CTkLabel(main_frame, text="替换为:", anchor="w").pack(fill="x", pady=(0, 5))
        self.replace_entry = ctk.CTkEntry(main_frame, width=400, height=35)
        self.replace_entry.pack(fill="x", pady=(0, 10))
        
        # 选项
        options_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        options_frame.pack(fill="x", pady=(0, 15))
        
        self.case_sensitive_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame, 
            text="区分大小写", 
            variable=self.case_sensitive_var,
            command=self._on_option_change
        ).pack(side="left", padx=(0, 15))
        
        self.regex_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame, 
            text="正则表达式", 
            variable=self.regex_var,
            command=self._on_option_change
        ).pack(side="left", padx=(0, 15))
        
        self.whole_word_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            options_frame, 
            text="全词匹配", 
            variable=self.whole_word_var,
            command=self._on_option_change
        ).pack(side="left")
        
        # 匹配计数
        self.match_label = ctk.CTkLabel(
            main_frame, 
            text="", 
            anchor="w",
            text_color=("gray50", "gray70")
        )
        self.match_label.pack(fill="x", pady=(0, 10))
        
        # 按钮区
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        # 上一个/下一个
        nav_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        nav_frame.pack(side="left")
        
        ctk.CTkButton(
            nav_frame, 
            text="◀ 上一个", 
            width=90,
            command=self.find_prev
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            nav_frame, 
            text="下一个 ▶", 
            width=90,
            command=self.find_next
        ).pack(side="left")
        
        # 替换按钮
        replace_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        replace_frame.pack(side="right")
        
        ctk.CTkButton(
            replace_frame, 
            text="替换", 
            width=80,
            command=self.replace_current
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            replace_frame, 
            text="全部替换", 
            width=100,
            fg_color=("green", "darkgreen"),
            hover_color=("darkgreen", "green"),
            command=self.replace_all
        ).pack(side="left")
        
        # 结果预览（可选）
        ctk.CTkLabel(
            main_frame, 
            text="匹配预览:", 
            anchor="w"
        ).pack(fill="x", pady=(20, 5))
        
        self.results_text = ctk.CTkTextbox(main_frame, height=100, wrap="none")
        self.results_text.pack(fill="both", expand=True)
        self.results_text.configure(state="disabled")
        
        # 关闭时清除高亮
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 设置焦点
        self.search_entry.focus()
        
        # 如果有选中文本，填入搜索框
        self._fill_selection()
    
    def _fill_selection(self):
        """将选中文本填入搜索框"""
        try:
            textbox = self._get_textbox()
            if textbox:
                try:
                    selected = textbox.get("sel.first", "sel.last")
                    if selected and len(selected) < 100:
                        self.search_entry.delete(0, END)
                        self.search_entry.insert(0, selected)
                except Exception:
                    pass
        except Exception:
            pass
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
    
    def _on_search_change(self, event=None):
        """搜索文本变化时实时高亮"""
        self._do_search()
    
    def _on_option_change(self):
        """选项变化时重新搜索"""
        self._do_search()
    
    def _do_search(self):
        """执行搜索并高亮"""
        textbox = self._get_textbox()
        if not textbox:
            return
        
        # 清除之前的高亮
        self._clear_highlights()
        
        pattern = self.search_entry.get()
        if not pattern:
            self.matches = []
            self.match_label.configure(text="")
            self._update_preview()
            return
        
        # 获取文本
        text = textbox.get("1.0", "end-1c")
        
        # 构建正则模式
        try:
            flags = 0 if self.case_sensitive_var.get() else re.IGNORECASE
            
            if self.regex_var.get():
                regex = pattern
            else:
                regex = re.escape(pattern)
            
            if self.whole_word_var.get():
                regex = r'\b' + regex + r'\b'
            
            compiled = re.compile(regex, flags)
            
            # 查找所有匹配
            self.matches = []
            for match in compiled.finditer(text):
                start = match.start()
                end = match.end()
                # 转换为 Text 索引
                start_idx = self._offset_to_index(text, start)
                end_idx = self._offset_to_index(text, end)
                self.matches.append((start_idx, end_idx))
            
            # 高亮所有匹配
            self._setup_tags(textbox)
            for start_idx, end_idx in self.matches:
                textbox.tag_add(self.highlight_tag, start_idx, end_idx)
            
            # 更新计数
            count = len(self.matches)
            if count > 0:
                self.match_label.configure(
                    text=f"找到 {count} 个匹配",
                    text_color=("green", "lightgreen")
                )
                self.current_match_index = 0
                self._highlight_current()
            else:
                self.match_label.configure(
                    text="未找到匹配",
                    text_color=("red", "lightcoral")
                )
            
            # 更新编辑器热力图
            if hasattr(self.app, 'input_editor'):
                self.app.input_editor.update_search_heatmap(self.matches)
            
            # 更新预览区搜索高亮
            if hasattr(self.app, 'preview'):
                self.app.preview.highlight_search_term(pattern, self.case_sensitive_var.get())
            
            self._update_preview()
            
        except re.error as e:
            self.match_label.configure(
                text=f"正则表达式错误: {e}",
                text_color=("red", "lightcoral")
            )
            self.matches = []
    
    def _offset_to_index(self, text: str, offset: int) -> str:
        """将字符偏移转换为 Text 控件索引"""
        line = 1
        col = 0
        for i, char in enumerate(text):
            if i == offset:
                break
            if char == '\n':
                line += 1
                col = 0
            else:
                col += 1
        return f"{line}.{col}"
    
    def _setup_tags(self, textbox):
        """设置高亮标签"""
        try:
            textbox.tag_configure(
                self.highlight_tag, 
                background="#FFFF00",
                foreground="#000000"
            )
            textbox.tag_configure(
                self.current_tag, 
                background="#FF8C00",
                foreground="#FFFFFF"
            )
        except Exception:
            pass
    
    def _clear_highlights(self):
        """清除所有高亮"""
        textbox = self._get_textbox()
        if textbox:
            try:
                textbox.tag_remove(self.highlight_tag, "1.0", END)
                textbox.tag_remove(self.current_tag, "1.0", END)
            except Exception:
                pass
    
    def _highlight_current(self):
        """高亮当前匹配"""
        textbox = self._get_textbox()
        if not textbox or not self.matches:
            return
        
        try:
            # 移除之前的当前匹配高亮
            textbox.tag_remove(self.current_tag, "1.0", END)
            
            # 高亮当前
            start, end = self.matches[self.current_match_index]
            textbox.tag_add(self.current_tag, start, end)
            
            # 滚动到当前匹配
            textbox.see(start)
            
            # 更新标签
            self.match_label.configure(
                text=f"第 {self.current_match_index + 1}/{len(self.matches)} 个匹配",
                text_color=("green", "lightgreen")
            )
        except Exception:
            pass
    
    def find_next(self):
        """查找下一个"""
        if not self.matches:
            self._do_search()
            return
        
        self.current_match_index = (self.current_match_index + 1) % len(self.matches)
        self._highlight_current()
    
    def find_prev(self):
        """查找上一个"""
        if not self.matches:
            self._do_search()
            return
        
        self.current_match_index = (self.current_match_index - 1) % len(self.matches)
        self._highlight_current()
    
    def replace_current(self):
        """替换当前匹配"""
        textbox = self._get_textbox()
        if not textbox or not self.matches:
            return
        
        try:
            start, end = self.matches[self.current_match_index]
            replacement = self.replace_entry.get()
            
            # 如果使用正则，处理替换中的组引用
            if self.regex_var.get():
                pattern = self.search_entry.get()
                flags = 0 if self.case_sensitive_var.get() else re.IGNORECASE
                matched_text = textbox.get(start, end)
                try:
                    replacement = re.sub(pattern, replacement, matched_text, flags=flags)
                except Exception:
                    pass
            
            # 执行替换
            textbox.delete(start, end)
            textbox.insert(start, replacement)
            
            # 重新搜索
            self._do_search()
            
        except Exception as e:
            print(f"替换失败: {e}")
    
    def replace_all(self):
        """替换所有匹配"""
        textbox = self._get_textbox()
        if not textbox or not self.matches:
            return
        
        try:
            pattern = self.search_entry.get()
            replacement = self.replace_entry.get()
            
            # 获取当前文本
            text = textbox.get("1.0", "end-1c")
            
            # 构建正则
            flags = 0 if self.case_sensitive_var.get() else re.IGNORECASE
            
            if self.regex_var.get():
                regex = pattern
            else:
                regex = re.escape(pattern)
            
            if self.whole_word_var.get():
                regex = r'\b' + regex + r'\b'
            
            # 执行替换
            new_text, count = re.subn(regex, replacement, text, flags=flags)
            
            if count > 0:
                # 更新文本
                textbox.delete("1.0", END)
                textbox.insert("1.0", new_text)
                
                self.match_label.configure(
                    text=f"已替换 {count} 处",
                    text_color=("green", "lightgreen")
                )
                
                self.matches = []
                self._clear_highlights()
                self._update_preview()
            
        except re.error as e:
            self.match_label.configure(
                text=f"正则表达式错误: {e}",
                text_color=("red", "lightcoral")
            )
        except Exception as e:
            print(f"全部替换失败: {e}")
    
    def _update_preview(self):
        """更新匹配预览"""
        try:
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", END)
            
            textbox = self._get_textbox()
            if not textbox or not self.matches:
                self.results_text.configure(state="disabled")
                return
            
            # 显示前10个匹配的上下文
            for i, (start, end) in enumerate(self.matches[:10]):
                try:
                    # 获取匹配行
                    line_start = start.split('.')[0] + ".0"
                    line_end = start.split('.')[0] + ".end"
                    line_text = textbox.get(line_start, line_end)
                    matched = textbox.get(start, end)
                    
                    # 显示行号和内容
                    line_num = start.split('.')[0]
                    self.results_text.insert(END, f"L{line_num}: ")
                    self.results_text.insert(END, line_text.strip()[:80])
                    if len(line_text.strip()) > 80:
                        self.results_text.insert(END, "...")
                    self.results_text.insert(END, "\n")
                except Exception:
                    pass
            
            if len(self.matches) > 10:
                self.results_text.insert(END, f"\n... 还有 {len(self.matches) - 10} 个匹配")
            
            self.results_text.configure(state="disabled")
        except Exception:
            pass
    
    def _on_close(self):
        """关闭对话框时清除高亮"""
        self._clear_highlights()
        # 清除热力图
        if hasattr(self.app, 'input_editor'):
            self.app.input_editor.update_search_heatmap([])
        # 清除预览区搜索高亮
        if hasattr(self.app, 'preview'):
            self.app.preview.highlight_search_term("")
        self.dialog.destroy()
        self.dialog = None
