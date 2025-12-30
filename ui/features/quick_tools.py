# -*- coding: utf-8 -*-
"""
快捷功能优化 - 快速格式化、智能目录、脚注等
"""

import re
import customtkinter as ctk
from tkinter import messagebox


class QuickToolsFeature:
    """快捷功能工具集"""
    
    def __init__(self, app):
        self.app = app
        
    def quick_format_document(self):
        """一键美化文档"""
        content = self.app.input_text.get("1.0", "end-1c")
        
        if not content.strip():
            messagebox.showwarning("提示", "文档内容为空！")
            return
            
        # 执行格式化
        formatted = self._format_document(content)
        
        # 应用到编辑器
        self.app.input_text.delete("1.0", "end")
        self.app.input_text.insert("1.0", formatted)
        self.app.on_text_change(None)
        
        messagebox.showinfo("完成", "文档格式化完成！")
        self.app.update_status("✅ 文档已美化")
        
    def _format_document(self, content: str) -> str:
        """格式化文档内容"""
        lines = content.split('\n')
        formatted_lines = []
        in_code_block = False
        
        for line in lines:
            # 检查代码块
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                formatted_lines.append(line)
                continue
                
            # 代码块内不处理
            if in_code_block:
                formatted_lines.append(line)
                continue
                
            # 规范化标题
            if line.startswith('#'):
                # 确保#后有空格
                line = re.sub(r'^(#+)([^\s])', r'\1 \2', line)
                
            # 规范化列表
            line = re.sub(r'^(\s*)([\*\-\+])([^\s])', r'\1\2 \3', line)
            line = re.sub(r'^(\s*)(\d+\.)([^\s])', r'\1\2 \3', line)
            
            # 规范化强调符号
            line = re.sub(r'\*\*([^\*]+)\*\*', r'**\1**', line)
            line = re.sub(r'\*([^\*]+)\*', r'*\1*', line)
            
            # 删除行尾空格
            line = line.rstrip()
            
            formatted_lines.append(line)
            
        # 确保文末有一个空行
        result = '\n'.join(formatted_lines)
        if not result.endswith('\n'):
            result += '\n'
            
        return result
        
    def insert_table_of_contents(self):
        """插入智能目录"""
        content = self.app.input_text.get("1.0", "end-1c")
        
        # 提取所有标题
        headings = self._extract_headings(content)
        
        if not headings:
            messagebox.showwarning("提示", "文档中没有找到标题！")
            return
            
        # 生成目录
        toc = self._generate_toc(headings)
        
        # 插入到文档开头
        self.app.input_text.insert("1.0", toc + "\n\n---\n\n")
        self.app.on_text_change(None)
        
        messagebox.showinfo("完成", f"已插入目录（包含 {len(headings)} 个标题）")
        self.app.update_status("✅ 目录已插入")
        
    def _extract_headings(self, content: str) -> list:
        """提取文档中的所有标题"""
        headings = []
        
        for line in content.split('\n'):
            match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()
                # 移除可能的锚点
                title = re.sub(r'\s*\{#[^\}]+\}\s*$', '', title)
                headings.append((level, title))
                
        return headings
        
    def _generate_toc(self, headings: list) -> str:
        """生成目录"""
        toc_lines = ["# 目录\n"]
        
        for level, title in headings:
            indent = "  " * (level - 1)
            # 创建锚点链接（简化版，实际应该处理特殊字符）
            anchor = title.lower().replace(' ', '-')
            anchor = re.sub(r'[^\w\-]', '', anchor)
            
            toc_lines.append(f"{indent}- [{title}](#{anchor})")
            
        return '\n'.join(toc_lines)
        
    def insert_footnote(self):
        """插入脚注"""
        dialog = ctk.CTkInputDialog(
            text="请输入脚注内容:",
            title="插入脚注"
        )
        footnote_text = dialog.get_input()
        
        if not footnote_text:
            return
            
        # 获取当前脚注数量
        content = self.app.input_text.get("1.0", "end-1c")
        footnote_count = len(re.findall(r'\[\^(\d+)\]', content))
        next_num = footnote_count + 1
        
        # 插入脚注引用
        self.app.input_text.insert("insert", f"[^{next_num}]")
        
        # 在文档末尾添加脚注定义
        self.app.input_text.insert("end", f"\n\n[^{next_num}]: {footnote_text}")
        
        self.app.on_text_change(None)
        self.app.update_status(f"✅ 已插入脚注 {next_num}")
        
    def insert_citation(self):
        """插入文献引用"""
        # 显示引用对话框
        citation_dialog = ctk.CTkToplevel(self.app)
        citation_dialog.title("📚 插入引用")
        citation_dialog.geometry("500x400")
        citation_dialog.transient(self.app)
        
        ctk.CTkLabel(
            citation_dialog,
            text="添加文献引用",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=15)
        
        # 引用格式选择
        format_frame = ctk.CTkFrame(citation_dialog, fg_color="transparent")
        format_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(format_frame, text="引用格式:").pack(anchor="w", pady=5)
        
        format_var = ctk.StringVar(value="APA")
        ctk.CTkSegmentedButton(
            format_frame,
            values=["APA", "MLA", "Chicago", "GB/T 7714"],
            variable=format_var
        ).pack(anchor="w", pady=5)
        
        # 引用信息输入
        info_frame = ctk.CTkFrame(citation_dialog, fg_color="transparent")
        info_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        fields = [
            ("作者", "author"),
            ("标题", "title"),
            ("年份", "year"),
            ("出版社/期刊", "publisher"),
            ("页码", "pages")
        ]
        
        entries = {}
        for label, key in fields:
            ctk.CTkLabel(info_frame, text=f"{label}:").pack(anchor="w", pady=2)
            entry = ctk.CTkEntry(info_frame, width=440)
            entry.pack(anchor="w", pady=2)
            entries[key] = entry
            
        # 插入按钮
        def insert_citation_text():
            # 简化的引用格式生成
            author = entries['author'].get()
            year = entries['year'].get()
            
            if not author or not year:
                messagebox.showwarning("提示", "至少需要填写作者和年份！")
                return
                
            citation_text = f"({author}, {year})"
            self.app.input_text.insert("insert", citation_text)
            
            # 生成完整引用（添加到文档末尾）
            full_citation = self._generate_full_citation(
                format_var.get(),
                {k: v.get() for k, v in entries.items()}
            )
            
            self.app.input_text.insert("end", f"\n\n{full_citation}")
            self.app.on_text_change(None)
            
            citation_dialog.destroy()
            self.app.update_status("✅ 引用已插入")
            
        ctk.CTkButton(
            citation_dialog,
            text="✅ 插入",
            command=insert_citation_text,
            width=120
        ).pack(pady=15)
        
    def _generate_full_citation(self, format_type: str, info: dict) -> str:
        """生成完整引用格式"""
        author = info.get('author', '')
        title = info.get('title', '')
        year = info.get('year', '')
        publisher = info.get('publisher', '')
        pages = info.get('pages', '')
        
        if format_type == "APA":
            citation = f"{author} ({year}). {title}."
            if publisher:
                citation += f" {publisher}."
            if pages:
                citation += f" pp. {pages}."
                
        elif format_type == "MLA":
            citation = f"{author}. \"{title}.\" {publisher}, {year}."
            if pages:
                citation += f" {pages}."
                
        elif format_type == "GB/T 7714":
            citation = f"{author}. {title}[M]. {publisher}, {year}"
            if pages:
                citation += f": {pages}"
            citation += "."
            
        else:  # Chicago
            citation = f"{author}. {year}. {title}. {publisher}."
            
        return citation
        
    def insert_cross_reference(self):
        """插入交叉引用"""
        content = self.app.input_text.get("1.0", "end-1c")
        
        # 提取所有可引用的元素
        headings = self._extract_headings(content)
        images = re.findall(r'!\[([^\]]+)\]', content)
        tables = re.findall(r'\|([^\|]+)\|', content)[:5]  # 简化，只取前几个
        
        # 创建引用对话框
        ref_dialog = ctk.CTkToplevel(self.app)
        ref_dialog.title("🔗 插入交叉引用")
        ref_dialog.geometry("500x400")
        ref_dialog.transient(self.app)
        
        ctk.CTkLabel(
            ref_dialog,
            text="选择要引用的元素",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=15)
        
        # 类型选择
        type_frame = ctk.CTkFrame(ref_dialog, fg_color="transparent")
        type_frame.pack(fill="x", padx=20, pady=10)
        
        ref_type_var = ctk.StringVar(value="heading")
        
        ctk.CTkRadioButton(
            type_frame,
            text=f"标题 ({len(headings)}个)",
            variable=ref_type_var,
            value="heading"
        ).pack(anchor="w", pady=3)
        
        ctk.CTkRadioButton(
            type_frame,
            text=f"图片 ({len(images)}个)",
            variable=ref_type_var,
            value="image"
        ).pack(anchor="w", pady=3)
        
        ctk.CTkRadioButton(
            type_frame,
            text=f"表格 ({len(tables)}个)",
            variable=ref_type_var,
            value="table"
        ).pack(anchor="w", pady=3)
        
        # 元素列表
        list_frame = ctk.CTkScrollableFrame(ref_dialog, width=460, height=180)
        list_frame.pack(padx=20, pady=10)
        
        # TODO: 显示可选元素列表
        
        ctk.CTkButton(
            ref_dialog,
            text="关闭",
            command=ref_dialog.destroy,
            width=100
        ).pack(pady=10)
        
    def optimize_heading_structure(self):
        """优化标题层级结构"""
        content = self.app.input_text.get("1.0", "end-1c")
        headings = self._extract_headings(content)
        
        if not headings:
            messagebox.showwarning("提示", "文档中没有标题！")
            return
            
        # 检查标题层级
        issues = []
        prev_level = 0
        
        for i, (level, title) in enumerate(headings):
            # 检查跳级
            if prev_level > 0 and level > prev_level + 1:
                issues.append(f"第{i+1}个标题 '{title}' 跳级（从H{prev_level}跳到H{level}）")
                
            prev_level = level
            
        if issues:
            msg = "发现标题结构问题:\n\n"
            msg += '\n'.join(issues)
            msg += "\n\n是否自动修复？"
            
            result = messagebox.askyesno("标题结构检查", msg)
            
            if result:
                # TODO: 实现自动修复逻辑
                messagebox.showinfo("提示", "自动修复功能开发中...")
        else:
            messagebox.showinfo("检查完成", "标题结构正常！✅")
