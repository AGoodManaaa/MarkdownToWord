# -*- coding: utf-8 -*-
"""
目录（TOC）自动生成功能模块
"""

import re
import customtkinter as ctk
from tkinter import messagebox
from typing import TYPE_CHECKING, List, Dict

if TYPE_CHECKING:
    from gui import App


class TOCGeneratorFeature:
    """目录自动生成功能"""
    
    def __init__(self, app: 'App'):
        self.app = app
    
    def extract_headings(self, markdown_text: str) -> List[Dict]:
        """
        提取Markdown文本中的所有标题
        
        Returns:
            List of {'level': int, 'text': str, 'anchor': str, 'line': int}
        """
        headings = []
        lines = markdown_text.split('\n')
        
        for line_num, line in enumerate(lines, 1):
            # 匹配 ATX 风格标题 (# 标题)
            match = re.match(r'^(#{1,6})\s+(.+)$', line.strip())
            if match:
                level = len(match.group(1))
                text = match.group(2).strip()
                
                # 生成锚点（GitHub风格）
                anchor = self._generate_anchor(text)
                
                headings.append({
                    'level': level,
                    'text': text,
                    'anchor': anchor,
                    'line': line_num
                })
        
        return headings
    
    def _generate_anchor(self, text: str) -> str:
        """
        生成GitHub风格的锚点
        
        例如：
        "第一章 介绍" -> "第一章-介绍"
        "1.1 Background" -> "11-background"
        """
        # 转换为小写
        anchor = text.lower()
        
        # 移除特殊字符，保留字母数字和中文
        anchor = re.sub(r'[^\w\s\u4e00-\u9fff-]', '', anchor)
        
        # 替换空格为连字符
        anchor = re.sub(r'\s+', '-', anchor.strip())
        
        return anchor
    
    def generate_toc(self, headings: List[Dict], style: str = 'bullet', 
                     min_level: int = 1, max_level: int = 6) -> str:
        """
        生成目录Markdown
        
        Args:
            headings: 标题列表
            style: 'bullet' 或 'numbered'
            min_level: 最小标题级别
            max_level: 最大标题级别
        
        Returns:
            Markdown格式的目录
        """
        if not headings:
            return "<!-- 未找到标题 -->"
        
        # 过滤标题级别
        filtered_headings = [
            h for h in headings 
            if min_level <= h['level'] <= max_level
        ]
        
        if not filtered_headings:
            return "<!-- 未找到符合条件的标题 -->"
        
        toc_lines = ["## 目录", ""]
        
        for heading in filtered_headings:
            indent = "  " * (heading['level'] - min_level)
            link_text = heading['text']
            link_anchor = heading['anchor']
            
            if style == 'numbered':
                # 有序列表
                toc_lines.append(f"{indent}1. [{link_text}](#{link_anchor})")
            else:
                # 无序列表
                toc_lines.append(f"{indent}- [{link_text}](#{link_anchor})")
        
        toc_lines.append("")  # 空行
        return "\n".join(toc_lines)
    
    def insert_toc_at_cursor(self, style: str = 'bullet'):
        """在光标位置插入目录"""
        try:
            # 获取当前文本
            text_widget = self.app.input_text._textbox
            current_text = text_widget.get("1.0", "end-1c")
            
            # 提取标题
            headings = self.extract_headings(current_text)
            
            if not headings:
                messagebox.showinfo("提示", "当前文档没有找到任何标题！\n\n请先添加标题（使用 # 符号）。")
                return
            
            # 生成目录
            toc = self.generate_toc(headings, style=style)
            
            # 在光标位置插入
            text_widget.insert("insert", toc + "\n\n")
            
            # 更新预览
            self.app.on_text_change(None)
            
            self.app.update_status(f"✅ 已插入目录（{len(headings)} 个标题）")
            
        except Exception as e:
            messagebox.showerror("错误", f"插入目录失败：{str(e)}")
    
    def show_toc_dialog(self):
        """显示目录配置对话框"""
        dialog = TOCDialog(self.app, self)
        dialog.grab_set()
        dialog.wait_window()


class TOCDialog(ctk.CTkToplevel):
    """目录生成配置对话框"""
    
    def __init__(self, parent, toc_feature: TOCGeneratorFeature):
        super().__init__(parent)
        
        self.toc_feature = toc_feature
        
        # 窗口配置
        self.title("插入目录")
        self.geometry("400x300")
        self.resizable(False, False)
        
        # 居中显示
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.winfo_screenheight() // 2) - (300 // 2)
        self.geometry(f"+{x}+{y}")
        
        self._create_ui()
    
    def _create_ui(self):
        """创建界面"""
        # 标题
        title_label = ctk.CTkLabel(
            self,
            text="📑 目录生成设置",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 样式选择
        style_frame = ctk.CTkFrame(self, fg_color="transparent")
        style_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(
            style_frame,
            text="目录样式：",
            font=ctk.CTkFont(size=14)
        ).pack(side="left", padx=10)
        
        self.style_var = ctk.StringVar(value="bullet")
        
        ctk.CTkRadioButton(
            style_frame,
            text="无序列表 (- )",
            variable=self.style_var,
            value="bullet"
        ).pack(side="left", padx=10)
        
        ctk.CTkRadioButton(
            style_frame,
            text="有序列表 (1. )",
            variable=self.style_var,
            value="numbered"
        ).pack(side="left", padx=10)
        
        # 预览
        preview_frame = ctk.CTkFrame(self)
        preview_frame.pack(pady=15, padx=20, fill="both", expand=True)
        
        ctk.CTkLabel(
            preview_frame,
            text="预览：",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))
        
        # 获取当前文档的标题进行预览
        self._update_preview(preview_frame)
        
        # 按钮
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=15, padx=20, fill="x")
        
        ctk.CTkButton(
            button_frame,
            text="插入目录",
            command=self._insert_toc,
            width=150
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            button_frame,
            text="取消",
            command=self.destroy,
            fg_color="gray",
            width=150
        ).pack(side="right", padx=5)
    
    def _update_preview(self, parent_frame):
        """更新预览"""
        try:
            # 获取当前文本
            text_widget = self.toc_feature.app.input_text._textbox
            current_text = text_widget.get("1.0", "end-1c")
            
            # 提取标题
            headings = self.toc_feature.extract_headings(current_text)
            
            preview_text = ctk.CTkTextbox(
                parent_frame,
                height=120,
                font=ctk.CTkFont(family="Consolas", size=11)
            )
            preview_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)
            
            if headings:
                # 生成预览（只显示前5个标题）
                preview_headings = headings[:5]
                toc_preview = self.toc_feature.generate_toc(
                    preview_headings, 
                    style=self.style_var.get()
                )
                preview_text.insert("1.0", toc_preview)
                
                if len(headings) > 5:
                    preview_text.insert("end", f"\n... 还有 {len(headings) - 5} 个标题")
            else:
                preview_text.insert("1.0", "当前文档没有找到标题")
            
            preview_text.configure(state="disabled")
            
        except Exception as e:
            print(f"预览更新失败: {e}")
    
    def _insert_toc(self):
        """插入目录"""
        style = self.style_var.get()
        self.toc_feature.insert_toc_at_cursor(style=style)
        self.destroy()
