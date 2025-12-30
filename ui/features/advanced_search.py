# -*- coding: utf-8 -*-
"""
智能搜索和替换增强功能
"""

import re
import customtkinter as ctk
from tkinter import messagebox


class AdvancedSearchFeature:
    """高级搜索功能"""
    
    def __init__(self, app):
        self.app = app
        self.search_dialog = None
        self.search_history = []
        self.current_matches = []
        self.current_match_index = -1
        
    def show_advanced_search_dialog(self):
        """显示高级搜索对话框"""
        if self.search_dialog and self.search_dialog.winfo_exists():
            self.search_dialog.focus()
            return
            
        self.search_dialog = ctk.CTkToplevel(self.app)
        self.search_dialog.title("🔍 高级搜索与替换")
        self.search_dialog.geometry("700x650")
        self.search_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.search_dialog,
            text="高级搜索与替换",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=15)
        
        # 搜索输入区域
        search_frame = ctk.CTkFrame(self.search_dialog, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(search_frame, text="搜索:").grid(row=0, column=0, sticky="w", pady=5)
        self.search_entry = ctk.CTkEntry(search_frame, width=500)
        self.search_entry.grid(row=0, column=1, padx=10, pady=5)
        self.search_entry.bind('<Return>', lambda e: self._search())
        
        ctk.CTkLabel(search_frame, text="替换为:").grid(row=1, column=0, sticky="w", pady=5)
        self.replace_entry = ctk.CTkEntry(search_frame, width=500)
        self.replace_entry.grid(row=1, column=1, padx=10, pady=5)
        
        # 搜索选项
        options_frame = ctk.CTkFrame(self.search_dialog)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # 左侧选项
        left_options = ctk.CTkFrame(options_frame, fg_color="transparent")
        left_options.pack(side="left", fill="both", expand=True)
        
        self.case_sensitive_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            left_options,
            text="区分大小写",
            variable=self.case_sensitive_var
        ).pack(anchor="w", pady=3)
        
        self.whole_word_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            left_options,
            text="全词匹配",
            variable=self.whole_word_var
        ).pack(anchor="w", pady=3)
        
        self.regex_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            left_options,
            text="正则表达式",
            variable=self.regex_var
        ).pack(anchor="w", pady=3)
        
        # 右侧选项
        right_options = ctk.CTkFrame(options_frame, fg_color="transparent")
        right_options.pack(side="right", fill="both", expand=True)
        
        self.highlight_all_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            right_options,
            text="高亮所有匹配项",
            variable=self.highlight_all_var
        ).pack(anchor="w", pady=3)
        
        # 搜索范围
        scope_frame = ctk.CTkFrame(right_options, fg_color="transparent")
        scope_frame.pack(anchor="w", pady=3)
        
        ctk.CTkLabel(scope_frame, text="搜索范围:").pack(side="left", padx=(0, 5))
        
        self.search_scope_var = ctk.StringVar(value="all")
        ctk.CTkRadioButton(
            scope_frame,
            text="全部",
            variable=self.search_scope_var,
            value="all"
        ).pack(side="left", padx=3)
        
        ctk.CTkRadioButton(
            scope_frame,
            text="选中区域",
            variable=self.search_scope_var,
            value="selection"
        ).pack(side="left", padx=3)
        
        # 元素类型过滤
        filter_frame = ctk.CTkFrame(self.search_dialog)
        filter_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(
            filter_frame,
            text="仅搜索特定元素:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", pady=5)
        
        element_types_frame = ctk.CTkFrame(filter_frame, fg_color="transparent")
        element_types_frame.pack(fill="x", pady=5)
        
        self.search_in_headings_var = ctk.BooleanVar(value=False)
        self.search_in_code_var = ctk.BooleanVar(value=False)
        self.search_in_links_var = ctk.BooleanVar(value=False)
        
        ctk.CTkCheckBox(
            element_types_frame,
            text="标题",
            variable=self.search_in_headings_var
        ).pack(side="left", padx=5)
        
        ctk.CTkCheckBox(
            element_types_frame,
            text="代码块",
            variable=self.search_in_code_var
        ).pack(side="left", padx=5)
        
        ctk.CTkCheckBox(
            element_types_frame,
            text="链接文本",
            variable=self.search_in_links_var
        ).pack(side="left", padx=5)
        
        # 结果显示
        results_frame = ctk.CTkFrame(self.search_dialog)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        result_label = ctk.CTkLabel(
            results_frame,
            text="搜索结果:",
            font=ctk.CTkFont(weight="bold")
        )
        result_label.pack(anchor="w", pady=5)
        
        self.results_text = ctk.CTkTextbox(results_frame, height=200)
        self.results_text.pack(fill="both", expand=True, pady=5)
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(self.search_dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        # 左侧搜索按钮组
        search_btns = ctk.CTkFrame(btn_frame, fg_color="transparent")
        search_btns.pack(side="left")
        
        ctk.CTkButton(
            search_btns,
            text="🔍 搜索",
            command=self._search,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            search_btns,
            text="⬇️ 下一个",
            command=self._find_next,
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            search_btns,
            text="⬆️ 上一个",
            command=self._find_previous,
            width=100
        ).pack(side="left", padx=3)
        
        # 右侧替换按钮组
        replace_btns = ctk.CTkFrame(btn_frame, fg_color="transparent")
        replace_btns.pack(side="right")
        
        ctk.CTkButton(
            replace_btns,
            text="替换",
            command=self._replace_current,
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            replace_btns,
            text="全部替换",
            command=self._replace_all,
            fg_color="#10B981",
            hover_color="#059669",
            width=100
        ).pack(side="left", padx=3)
        
        # 搜索历史
        self._show_search_history()
        
    def _search(self):
        """执行搜索"""
        query = self.search_entry.get()
        if not query:
            messagebox.showwarning("提示", "请输入搜索内容！")
            return
            
        # 添加到历史
        if query not in self.search_history:
            self.search_history.insert(0, query)
            if len(self.search_history) > 10:
                self.search_history = self.search_history[:10]
                
        # 获取搜索内容
        text_widget = self.app.input_text._textbox
        
        if self.search_scope_var.get() == "selection":
            try:
                content = text_widget.get("sel.first", "sel.last")
                start_pos = "sel.first"
            except:
                content = text_widget.get("1.0", "end-1c")
                start_pos = "1.0"
        else:
            content = text_widget.get("1.0", "end-1c")
            start_pos = "1.0"
            
        # 元素过滤
        if any([self.search_in_headings_var.get(), 
                self.search_in_code_var.get(), 
                self.search_in_links_var.get()]):
            content = self._filter_by_elements(content)
            
        # 执行搜索
        self.current_matches = []
        
        flags = 0 if self.case_sensitive_var.get() else re.IGNORECASE
        
        if self.regex_var.get():
            # 正则搜索
            try:
                pattern = re.compile(query, flags)
                for match in pattern.finditer(content):
                    self.current_matches.append((match.start(), match.end(), match.group()))
            except re.error as e:
                messagebox.showerror("正则表达式错误", f"无效的正则表达式:\n{e}")
                return
        else:
            # 普通搜索
            if self.whole_word_var.get():
                pattern = re.compile(r'\b' + re.escape(query) + r'\b', flags)
                for match in pattern.finditer(content):
                    self.current_matches.append((match.start(), match.end(), match.group()))
            else:
                # 简单字符串搜索
                search_text = content if self.case_sensitive_var.get() else content.lower()
                search_query = query if self.case_sensitive_var.get() else query.lower()
                
                pos = 0
                while True:
                    pos = search_text.find(search_query, pos)
                    if pos == -1:
                        break
                    self.current_matches.append((pos, pos + len(query), content[pos:pos + len(query)]))
                    pos += 1
                    
        # 显示结果
        self._display_search_results()
        
        # 高亮所有匹配项
        if self.highlight_all_var.get():
            self._highlight_matches()
            
        # 跳转到第一个匹配项
        if self.current_matches:
            self.current_match_index = 0
            self._jump_to_match(0)
            
    def _filter_by_elements(self, content: str) -> str:
        """根据元素类型过滤内容"""
        filtered_parts = []
        
        if self.search_in_headings_var.get():
            # 提取标题
            headings = re.findall(r'^#+\s+.*$', content, re.MULTILINE)
            filtered_parts.extend(headings)
            
        if self.search_in_code_var.get():
            # 提取代码块
            code_blocks = re.findall(r'```[\s\S]*?```', content)
            inline_code = re.findall(r'`[^`]+`', content)
            filtered_parts.extend(code_blocks + inline_code)
            
        if self.search_in_links_var.get():
            # 提取链接文本
            links = re.findall(r'\[([^\]]+)\]', content)
            filtered_parts.extend(links)
            
        return '\n'.join(filtered_parts) if filtered_parts else content
        
    def _display_search_results(self):
        """显示搜索结果"""
        self.results_text.delete("1.0", "end")
        
        if self.current_matches:
            result = f"找到 {len(self.current_matches)} 个匹配项:\n\n"
            
            for i, (start, end, text) in enumerate(self.current_matches[:50], 1):  # 最多显示前50个
                result += f"{i}. {text}\n"
                
            if len(self.current_matches) > 50:
                result += f"\n... 还有 {len(self.current_matches) - 50} 个匹配项未显示"
                
            self.results_text.insert("1.0", result)
        else:
            self.results_text.insert("1.0", "未找到匹配项")
            
    def _highlight_matches(self):
        """高亮所有匹配项"""
        text_widget = self.app.input_text._textbox
        
        # 清除之前的高亮
        text_widget.tag_remove("search_highlight", "1.0", "end")
        
        # 配置高亮样式
        text_widget.tag_config("search_highlight", background="#FFEB3B", foreground="#000000")
        
        # 添加高亮
        for start, end, _ in self.current_matches:
            start_index = f"1.0 + {start} chars"
            end_index = f"1.0 + {end} chars"
            text_widget.tag_add("search_highlight", start_index, end_index)
            
    def _jump_to_match(self, index: int):
        """跳转到指定匹配项"""
        if 0 <= index < len(self.current_matches):
            text_widget = self.app.input_text._textbox
            start, end, _ = self.current_matches[index]
            
            start_index = f"1.0 + {start} chars"
            end_index = f"1.0 + {end} chars"
            
            text_widget.see(start_index)
            text_widget.tag_remove("sel", "1.0", "end")
            text_widget.tag_add("sel", start_index, end_index)
            text_widget.mark_set("insert", start_index)
            
    def _find_next(self):
        """查找下一个"""
        if not self.current_matches:
            self._search()
            return
            
        self.current_match_index = (self.current_match_index + 1) % len(self.current_matches)
        self._jump_to_match(self.current_match_index)
        
    def _find_previous(self):
        """查找上一个"""
        if not self.current_matches:
            self._search()
            return
            
        self.current_match_index = (self.current_match_index - 1) % len(self.current_matches)
        self._jump_to_match(self.current_match_index)
        
    def _replace_current(self):
        """替换当前匹配项"""
        if not self.current_matches or self.current_match_index < 0:
            messagebox.showwarning("提示", "请先执行搜索！")
            return
            
        replace_text = self.replace_entry.get()
        text_widget = self.app.input_text._textbox
        
        start, end, _ = self.current_matches[self.current_match_index]
        start_index = f"1.0 + {start} chars"
        end_index = f"1.0 + {end} chars"
        
        text_widget.delete(start_index, end_index)
        text_widget.insert(start_index, replace_text)
        
        # 更新匹配列表
        self.current_matches.pop(self.current_match_index)
        
        # 跳转到下一个
        if self.current_matches:
            self._find_next()
        else:
            messagebox.showinfo("完成", "所有匹配项已处理")
            
        self.app.on_text_change(None)
        
    def _replace_all(self):
        """替换所有匹配项"""
        if not self.current_matches:
            messagebox.showwarning("提示", "请先执行搜索！")
            return
            
        replace_text = self.replace_entry.get()
        count = len(self.current_matches)
        
        result = messagebox.askyesno(
            "确认",
            f"确定要替换所有 {count} 个匹配项吗？"
        )
        
        if not result:
            return
            
        text_widget = self.app.input_text._textbox
        
        # 从后往前替换，避免位置偏移
        for start, end, _ in reversed(self.current_matches):
            start_index = f"1.0 + {start} chars"
            end_index = f"1.0 + {end} chars"
            text_widget.delete(start_index, end_index)
            text_widget.insert(start_index, replace_text)
            
        self.current_matches = []
        self.app.on_text_change(None)
        
        messagebox.showinfo("完成", f"已替换 {count} 个匹配项")
        
    def _show_search_history(self):
        """显示搜索历史"""
        if not self.search_history:
            return
            
        # TODO: 可以添加一个下拉菜单显示历史记录
