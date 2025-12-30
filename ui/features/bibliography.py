# -*- coding: utf-8 -*-
"""
文献引用管理功能
支持 BibTeX 导入，[@cite] 语法解析，自动生成参考文献列表
"""

import re
import os
import json
import customtkinter as ctk
from tkinter import messagebox, END, filedialog
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass, asdict, field
from ui.dialog_utils import set_dialog_icon


@dataclass
class BibEntry:
    """文献条目"""
    key: str
    entry_type: str  # article, book, inproceedings, etc.
    title: str = ""
    author: str = ""
    year: str = ""
    journal: str = ""
    booktitle: str = ""
    publisher: str = ""
    volume: str = ""
    pages: str = ""
    doi: str = ""
    url: str = ""
    abstract: str = ""
    
    def to_citation(self, style: str = "apa") -> str:
        """生成引用格式文本"""
        if style == "apa":
            # APA 格式: Author (Year). Title. Journal, Volume, Pages.
            parts = []
            if self.author:
                parts.append(self.author)
            if self.year:
                parts.append(f"({self.year})")
            if self.title:
                parts.append(f"*{self.title}*")
            if self.journal:
                parts.append(f"{self.journal}")
            if self.volume:
                vol = self.volume
                if self.pages:
                    vol += f", {self.pages}"
                parts.append(vol)
            if self.doi:
                parts.append(f"https://doi.org/{self.doi}")
            return ". ".join(parts) + "."
        elif style == "gb":
            # GB/T 7714 格式（中文标准）
            parts = []
            if self.author:
                parts.append(self.author)
            if self.title:
                parts.append(self.title)
            if self.journal:
                parts.append(f"[J]. {self.journal}")
            elif self.booktitle:
                parts.append(f"[C]// {self.booktitle}")
            elif self.publisher:
                parts.append(f"[M]. {self.publisher}")
            if self.year:
                parts.append(self.year)
            if self.volume:
                parts.append(f"{self.volume}")
            if self.pages:
                parts.append(f": {self.pages}")
            return ", ".join(parts) + "."
        else:
            # 简单格式
            return f"{self.author} ({self.year}). {self.title}."


class BibliographyFeature:
    """文献引用管理功能"""
    
    CITATION_STYLES = [
        ("apa", "APA (美国心理学会)"),
        ("gb", "GB/T 7714 (中国国标)"),
        ("simple", "简单格式"),
    ]
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.entries: Dict[str, BibEntry] = {}  # key -> entry
        self.config_file = os.path.join(os.path.dirname(__file__), '..', '..', 'bibliography.json')
        self._load_entries()
    
    def _load_entries(self):
        """加载保存的文献条目"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.entries = {k: BibEntry(**v) for k, v in data.items()}
        except Exception as e:
            print(f"加载文献库失败: {e}")
    
    def _save_entries(self):
        """保存文献条目"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                data = {k: asdict(v) for k, v in self.entries.items()}
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存文献库失败: {e}")
    
    def show_dialog(self):
        """显示文献管理对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📑 文献引用管理")
        self.dialog.geometry("800x600")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 800) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 600) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 顶部工具栏
        toolbar = ctk.CTkFrame(main_frame, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 10))
        
        ctk.CTkButton(
            toolbar, text="📥 导入 BibTeX", width=120,
            command=self._import_bibtex
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="➕ 手动添加", width=100,
            command=self._add_entry
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="📋 生成参考文献", width=120,
            fg_color=("green", "darkgreen"),
            command=self._generate_references
        ).pack(side="left")
        
        # 引用格式选择
        style_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        style_frame.pack(side="right")
        
        ctk.CTkLabel(style_frame, text="格式:").pack(side="left", padx=(0, 5))
        self.style_var = ctk.StringVar(value="apa")
        style_names = [name for _, name in self.CITATION_STYLES]
        style_values = [code for code, _ in self.CITATION_STYLES]
        ctk.CTkOptionMenu(
            style_frame, values=style_names, width=150,
            command=lambda x: self.style_var.set(style_values[style_names.index(x)])
        ).pack(side="left")
        
        # 文献列表
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        ctk.CTkLabel(list_frame, text=f"文献库 ({len(self.entries)} 条)", font=("", 13, "bold")).pack(anchor="w", padx=5, pady=5)
        
        self.list_container = ctk.CTkScrollableFrame(list_frame)
        self.list_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        self._refresh_list()
        
        # 底部说明
        ctk.CTkLabel(
            main_frame,
            text="💡 在文档中使用 [@key] 引用文献，然后点击「生成参考文献」插入引用列表",
            font=("", 11),
            text_color=("gray50", "gray70")
        ).pack(fill="x")
    
    def _refresh_list(self):
        """刷新文献列表"""
        for widget in self.list_container.winfo_children():
            widget.destroy()
        
        for key, entry in self.entries.items():
            item_frame = ctk.CTkFrame(self.list_container, fg_color=("gray90", "gray20"))
            item_frame.pack(fill="x", pady=2)
            
            # 信息
            info = f"[@{key}] {entry.author} ({entry.year})\n{entry.title}"
            ctk.CTkLabel(
                item_frame, text=info, anchor="w", justify="left"
            ).pack(side="left", fill="x", expand=True, padx=10, pady=8)
            
            # 操作按钮
            btn_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
            btn_frame.pack(side="right", padx=5)
            
            ctk.CTkButton(
                btn_frame, text="📋", width=30,
                command=lambda k=key: self._insert_citation(k)
            ).pack(side="left", padx=2)
            
            ctk.CTkButton(
                btn_frame, text="🗑️", width=30,
                fg_color=("red", "darkred"),
                command=lambda k=key: self._delete_entry(k)
            ).pack(side="left", padx=2)
    
    def _import_bibtex(self):
        """导入 BibTeX 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 BibTeX 文件",
            filetypes=[("BibTeX文件", "*.bib"), ("所有文件", "*.*")],
            parent=self.dialog
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            entries = self._parse_bibtex(content)
            for key, entry in entries.items():
                self.entries[key] = entry
            
            self._save_entries()
            self._refresh_list()
            messagebox.showinfo("成功", f"导入了 {len(entries)} 条文献")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")
    
    def _parse_bibtex(self, content: str) -> Dict[str, BibEntry]:
        """解析 BibTeX 内容"""
        entries = {}
        
        # 匹配 @type{key, ... }
        pattern = r'@(\w+)\s*\{\s*([^,]+)\s*,([^@]*)\}'
        
        for match in re.finditer(pattern, content, re.DOTALL):
            entry_type = match.group(1).lower()
            key = match.group(2).strip()
            fields_str = match.group(3)
            
            # 解析字段
            fields = {}
            field_pattern = r'(\w+)\s*=\s*[{"]([^}"]*)[}"]'
            for field_match in re.finditer(field_pattern, fields_str):
                field_name = field_match.group(1).lower()
                field_value = field_match.group(2).strip()
                fields[field_name] = field_value
            
            entries[key] = BibEntry(
                key=key,
                entry_type=entry_type,
                title=fields.get('title', ''),
                author=fields.get('author', ''),
                year=fields.get('year', ''),
                journal=fields.get('journal', ''),
                booktitle=fields.get('booktitle', ''),
                publisher=fields.get('publisher', ''),
                volume=fields.get('volume', ''),
                pages=fields.get('pages', ''),
                doi=fields.get('doi', ''),
                url=fields.get('url', ''),
                abstract=fields.get('abstract', ''),
            )
        
        return entries
    
    def _add_entry(self):
        """手动添加文献"""
        editor = ctk.CTkToplevel(self.dialog)
        editor.title("添加文献")
        editor.geometry("500x500")
        editor.transient(self.dialog)
        editor.grab_set()
        
        frame = ctk.CTkScrollableFrame(editor)
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        entries = {}
        fields = [
            ("key", "引用键 (必填)", "example2024"),
            ("title", "标题", ""),
            ("author", "作者", ""),
            ("year", "年份", "2024"),
            ("journal", "期刊", ""),
            ("publisher", "出版社", ""),
            ("volume", "卷号", ""),
            ("pages", "页码", ""),
            ("doi", "DOI", ""),
        ]
        
        for field, label, default in fields:
            ctk.CTkLabel(frame, text=f"{label}:").pack(anchor="w")
            entry = ctk.CTkEntry(frame, width=400)
            entry.pack(fill="x", pady=(0, 10))
            if default:
                entry.insert(0, default)
            entries[field] = entry
        
        def save():
            key = entries["key"].get().strip()
            if not key:
                messagebox.showwarning("警告", "请输入引用键")
                return
            
            self.entries[key] = BibEntry(
                key=key,
                entry_type="article",
                title=entries["title"].get(),
                author=entries["author"].get(),
                year=entries["year"].get(),
                journal=entries["journal"].get(),
                publisher=entries["publisher"].get(),
                volume=entries["volume"].get(),
                pages=entries["pages"].get(),
                doi=entries["doi"].get(),
            )
            
            self._save_entries()
            self._refresh_list()
            editor.destroy()
        
        ctk.CTkButton(frame, text="保存", command=save).pack(pady=15)
    
    def _delete_entry(self, key: str):
        """删除文献"""
        if messagebox.askyesno("确认", f"确定删除文献 [@{key}]？"):
            del self.entries[key]
            self._save_entries()
            self._refresh_list()
    
    def _insert_citation(self, key: str):
        """插入引用到文档"""
        textbox = self._get_textbox()
        if textbox:
            textbox.insert("insert", f"[@{key}]")
    
    def _generate_references(self):
        """生成参考文献列表"""
        textbox = self._get_textbox()
        if not textbox:
            return
        
        # 从文档中提取所有引用
        text = textbox.get("1.0", "end-1c")
        citations = re.findall(r'\[@([^\]]+)\]', text)
        
        if not citations:
            messagebox.showinfo("提示", "文档中没有找到引用 [@key]")
            return
        
        # 去重并按顺序
        seen = set()
        unique_citations = []
        for c in citations:
            if c not in seen:
                seen.add(c)
                unique_citations.append(c)
        
        # 生成参考文献列表
        style = self.style_var.get()
        lines = ["\n\n## 参考文献\n"]
        
        for i, key in enumerate(unique_citations, 1):
            if key in self.entries:
                entry = self.entries[key]
                citation = entry.to_citation(style)
                lines.append(f"[{i}] {citation}\n")
            else:
                lines.append(f"[{i}] ⚠️ 未找到文献: @{key}\n")
        
        # 插入到文档末尾
        ref_text = "\n".join(lines)
        textbox.insert("end", ref_text)
        
        self.dialog.destroy()
        self.dialog = None
        messagebox.showinfo("成功", f"已生成 {len(unique_citations)} 条参考文献")
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
