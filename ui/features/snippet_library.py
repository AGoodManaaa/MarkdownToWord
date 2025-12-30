# -*- coding: utf-8 -*-
"""
Markdown片段库功能
保存常用的代码块、表格模板等，支持快捷插入、分类管理、导入导出
"""

import os
import json
import customtkinter as ctk
from tkinter import END, filedialog, messagebox
from typing import List, Dict, Optional
from dataclasses import dataclass, asdict
from datetime import datetime
from ui.dialog_utils import set_dialog_icon


@dataclass
class Snippet:
    """片段数据结构"""
    id: str
    name: str
    content: str
    category: str = "通用"
    description: str = ""
    shortcut: str = ""
    created_at: str = ""
    used_count: int = 0


class SnippetLibraryFeature:
    """Markdown片段库功能"""
    
    DEFAULT_SNIPPETS = [
        Snippet(
            id="table_2x3",
            name="2x3 表格",
            content="| 列1 | 列2 | 列3 |\n|-----|-----|-----|\n| 内容 | 内容 | 内容 |",
            category="表格",
            description="2行3列的基础表格"
        ),
        Snippet(
            id="code_python",
            name="Python代码块",
            content="```python\n# 在此输入代码\n\n```",
            category="代码",
            description="Python代码块模板"
        ),
        Snippet(
            id="code_javascript",
            name="JavaScript代码块",
            content="```javascript\n// 在此输入代码\n\n```",
            category="代码",
            description="JavaScript代码块模板"
        ),
        Snippet(
            id="note_tip",
            name="提示框",
            content="> 💡 **提示**: 在此输入提示内容",
            category="通用",
            description="带图标的提示引用块"
        ),
        Snippet(
            id="note_warning",
            name="警告框",
            content="> ⚠️ **警告**: 在此输入警告内容",
            category="通用",
            description="带图标的警告引用块"
        ),
        Snippet(
            id="checkbox_list",
            name="任务列表",
            content="- [ ] 任务1\n- [ ] 任务2\n- [x] 已完成任务",
            category="列表",
            description="带复选框的任务列表"
        ),
        Snippet(
            id="footnote",
            name="脚注",
            content="这是带脚注的文字[^1]\n\n[^1]: 脚注内容",
            category="通用",
            description="脚注模板"
        ),
        Snippet(
            id="image",
            name="图片",
            content="![图片描述](图片路径)",
            category="媒体",
            description="图片引用模板"
        ),
        Snippet(
            id="link",
            name="链接",
            content="[链接文字](URL)",
            category="媒体",
            description="超链接模板"
        ),
        Snippet(
            id="math_inline",
            name="行内公式",
            content="$E = mc^2$",
            category="公式",
            description="行内数学公式"
        ),
        Snippet(
            id="math_block",
            name="公式块",
            content="$$\n\\sum_{i=1}^{n} x_i = x_1 + x_2 + \\cdots + x_n\n$$",
            category="公式",
            description="块级数学公式"
        ),
    ]
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.snippets: List[Snippet] = []
        self.config_file = os.path.join(os.path.dirname(__file__), '..', '..', 'snippets.json')
        self._load_snippets()
    
    def _load_snippets(self):
        """加载片段库"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.snippets = [Snippet(**s) for s in data]
            else:
                # 使用默认片段
                self.snippets = list(self.DEFAULT_SNIPPETS)
                self._save_snippets()
        except Exception as e:
            print(f"加载片段库失败: {e}")
            self.snippets = list(self.DEFAULT_SNIPPETS)
    
    def _save_snippets(self):
        """保存片段库"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                data = [asdict(s) for s in self.snippets]
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存片段库失败: {e}")
    
    def show_dialog(self):
        """显示片段库对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📝 Markdown 片段库")
        self.dialog.geometry("700x550")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 700) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 550) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 顶部工具栏
        toolbar = ctk.CTkFrame(main_frame, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 10))
        
        ctk.CTkButton(
            toolbar, text="➕ 新建", width=80,
            command=self._new_snippet
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="📥 导入", width=80,
            command=self._import_snippets
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="📤 导出", width=80,
            command=self._export_snippets
        ).pack(side="left")
        
        # 分类过滤
        self.category_var = ctk.StringVar(value="全部")
        categories = ["全部"] + list(set(s.category for s in self.snippets))
        ctk.CTkOptionMenu(
            toolbar,
            values=categories,
            variable=self.category_var,
            width=100,
            command=self._filter_by_category
        ).pack(side="right")
        
        ctk.CTkLabel(toolbar, text="分类:").pack(side="right", padx=(0, 5))
        
        # 左右分栏
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # 左侧：片段列表
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkLabel(left_frame, text="片段列表", font=("", 13, "bold")).pack(pady=5)
        
        self.snippets_listbox = ctk.CTkScrollableFrame(left_frame)
        self.snippets_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 右侧：预览和操作
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        
        ctk.CTkLabel(right_frame, text="片段预览", font=("", 13, "bold")).pack(pady=5)
        
        self.preview_text = ctk.CTkTextbox(right_frame, height=200)
        self.preview_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 操作按钮
        btn_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=5, pady=10)
        
        ctk.CTkButton(
            btn_frame, text="📋 插入", width=100,
            fg_color=("green", "darkgreen"),
            command=self._insert_selected
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="✏️ 编辑", width=80,
            command=self._edit_selected
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="🗑️ 删除", width=80,
            fg_color=("red", "darkred"),
            command=self._delete_selected
        ).pack(side="left")
        
        # 初始化列表
        self._refresh_list()
        
        self.selected_snippet: Optional[Snippet] = None
    
    def _refresh_list(self):
        """刷新片段列表"""
        # 清空现有列表
        for widget in self.snippets_listbox.winfo_children():
            widget.destroy()
        
        category_filter = self.category_var.get()
        
        for snippet in self.snippets:
            if category_filter != "全部" and snippet.category != category_filter:
                continue
            
            btn = ctk.CTkButton(
                self.snippets_listbox,
                text=f"{snippet.name}\n[{snippet.category}]",
                anchor="w",
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray80", "gray30"),
                height=50,
                command=lambda s=snippet: self._select_snippet(s)
            )
            btn.pack(fill="x", pady=2)
    
    def _select_snippet(self, snippet: Snippet):
        """选择片段"""
        self.selected_snippet = snippet
        self.preview_text.delete("1.0", END)
        self.preview_text.insert("1.0", snippet.content)
    
    def _filter_by_category(self, category: str):
        """按分类过滤"""
        self._refresh_list()
    
    def _insert_selected(self):
        """插入选中的片段"""
        if not self.selected_snippet:
            messagebox.showinfo("提示", "请先选择一个片段")
            return
        
        try:
            textbox = self._get_textbox()
            if textbox:
                textbox.insert("insert", self.selected_snippet.content)
                self.selected_snippet.used_count += 1
                self._save_snippets()
                self.dialog.destroy()
                self.dialog = None
        except Exception as e:
            messagebox.showerror("错误", f"插入失败: {e}")
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
    
    def _new_snippet(self):
        """新建片段"""
        self._show_snippet_editor()
    
    def _edit_selected(self):
        """编辑选中的片段"""
        if not self.selected_snippet:
            messagebox.showinfo("提示", "请先选择一个片段")
            return
        self._show_snippet_editor(self.selected_snippet)
    
    def _show_snippet_editor(self, snippet: Optional[Snippet] = None):
        """显示片段编辑器"""
        editor = ctk.CTkToplevel(self.dialog)
        editor.title("编辑片段" if snippet else "新建片段")
        editor.geometry("500x450")
        editor.transient(self.dialog)
        editor.grab_set()
        
        frame = ctk.CTkFrame(editor)
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 名称
        ctk.CTkLabel(frame, text="名称:").pack(anchor="w")
        name_entry = ctk.CTkEntry(frame, width=400)
        name_entry.pack(fill="x", pady=(0, 10))
        if snippet:
            name_entry.insert(0, snippet.name)
        
        # 分类
        ctk.CTkLabel(frame, text="分类:").pack(anchor="w")
        category_entry = ctk.CTkEntry(frame, width=400)
        category_entry.pack(fill="x", pady=(0, 10))
        if snippet:
            category_entry.insert(0, snippet.category)
        else:
            category_entry.insert(0, "通用")
        
        # 内容
        ctk.CTkLabel(frame, text="内容:").pack(anchor="w")
        content_text = ctk.CTkTextbox(frame, height=200)
        content_text.pack(fill="both", expand=True, pady=(0, 10))
        if snippet:
            content_text.insert("1.0", snippet.content)
        
        def save():
            name = name_entry.get().strip()
            category = category_entry.get().strip() or "通用"
            content = content_text.get("1.0", "end-1c")
            
            if not name:
                messagebox.showwarning("警告", "请输入名称")
                return
            if not content:
                messagebox.showwarning("警告", "请输入内容")
                return
            
            if snippet:
                snippet.name = name
                snippet.category = category
                snippet.content = content
            else:
                new_snippet = Snippet(
                    id=f"custom_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    name=name,
                    category=category,
                    content=content,
                    created_at=datetime.now().isoformat()
                )
                self.snippets.append(new_snippet)
            
            self._save_snippets()
            self._refresh_list()
            editor.destroy()
        
        ctk.CTkButton(frame, text="保存", command=save).pack(pady=10)
    
    def _delete_selected(self):
        """删除选中的片段"""
        if not self.selected_snippet:
            messagebox.showinfo("提示", "请先选择一个片段")
            return
        
        if messagebox.askyesno("确认", f"确定删除片段 '{self.selected_snippet.name}'？"):
            self.snippets = [s for s in self.snippets if s.id != self.selected_snippet.id]
            self.selected_snippet = None
            self._save_snippets()
            self._refresh_list()
            self.preview_text.delete("1.0", END)
    
    def _import_snippets(self):
        """导入片段"""
        file_path = filedialog.askopenfilename(
            title="导入片段",
            filetypes=[("JSON文件", "*.json")],
            parent=self.dialog
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                imported = [Snippet(**s) for s in data]
                
            # 合并（避免重复ID）
            existing_ids = {s.id for s in self.snippets}
            for s in imported:
                if s.id not in existing_ids:
                    self.snippets.append(s)
            
            self._save_snippets()
            self._refresh_list()
            messagebox.showinfo("成功", f"导入了 {len(imported)} 个片段")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")
    
    def _export_snippets(self):
        """导出片段"""
        file_path = filedialog.asksaveasfilename(
            title="导出片段",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")],
            parent=self.dialog
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                data = [asdict(s) for s in self.snippets]
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", f"已导出 {len(self.snippets)} 个片段")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")
