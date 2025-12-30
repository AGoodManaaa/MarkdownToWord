# -*- coding: utf-8 -*-
"""
思维导图转换功能
将 Markdown 列表转换为思维导图，支持生成 Mermaid 语法和图片
"""

import re
import os
import tempfile
import customtkinter as ctk
from tkinter import messagebox, END, filedialog
from typing import List, Optional, Tuple
from dataclasses import dataclass
from ui.dialog_utils import set_dialog_icon


@dataclass
class MindmapNode:
    """思维导图节点"""
    text: str
    level: int
    children: List['MindmapNode'] = None
    
    def __post_init__(self):
        if self.children is None:
            self.children = []


class MindmapFeature:
    """思维导图转换功能"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
    
    def show_dialog(self):
        """显示思维导图对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("🧠 思维导图转换")
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
        
        # 说明
        ctk.CTkLabel(
            main_frame,
            text="输入 Markdown 列表，自动转换为思维导图",
            font=("", 14)
        ).pack(pady=(0, 10))
        
        # 左右分栏
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # 左侧：输入
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkLabel(left_frame, text="Markdown 列表输入:").pack(anchor="w", padx=5, pady=5)
        
        self.input_text = ctk.CTkTextbox(left_frame, wrap="word")
        self.input_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 示例内容
        sample = """- 项目管理
  - 计划阶段
    - 需求分析
    - 可行性研究
    - 资源分配
  - 执行阶段
    - 开发
    - 测试
    - 部署
  - 收尾阶段
    - 验收
    - 文档整理
    - 项目总结"""
        self.input_text.insert("1.0", sample)
        
        # 右侧：预览
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        
        ctk.CTkLabel(right_frame, text="Mermaid 语法预览:").pack(anchor="w", padx=5, pady=5)
        
        self.output_text = ctk.CTkTextbox(right_frame, wrap="word")
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(
            btn_frame, text="🔄 转换预览", width=100,
            command=self._preview
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="📥 插入Mermaid代码", width=140,
            fg_color=("green", "darkgreen"),
            command=self._insert_mermaid
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="📋 复制代码", width=100,
            command=self._copy_code
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="📄 从选中文本", width=100,
            command=self._load_from_selection
        ).pack(side="right")
    
    def _parse_list(self, text: str) -> Optional[MindmapNode]:
        """解析 Markdown 列表为树结构"""
        lines = text.strip().split('\n')
        if not lines:
            return None
        
        root = None
        stack = []  # (node, level)
        
        for line in lines:
            # 匹配列表项
            match = re.match(r'^(\s*)([-*+]|\d+\.)\s+(.+)', line)
            if not match:
                continue
            
            indent = len(match.group(1))
            content = match.group(3).strip()
            
            # 计算层级（每2个空格一级）
            level = indent // 2
            
            node = MindmapNode(text=content, level=level)
            
            if root is None:
                root = node
                stack = [(node, level)]
            else:
                # 找到父节点
                while stack and stack[-1][1] >= level:
                    stack.pop()
                
                if stack:
                    parent = stack[-1][0]
                    parent.children.append(node)
                else:
                    # 没有父节点，创建虚拟根
                    if root.text != "__root__":
                        old_root = root
                        root = MindmapNode(text="__root__", level=-1)
                        root.children.append(old_root)
                    root.children.append(node)
                
                stack.append((node, level))
        
        return root
    
    def _node_to_mermaid(self, node: MindmapNode, depth: int = 0) -> str:
        """将节点转换为 Mermaid mindmap 语法"""
        lines = []
        indent = "  " * depth
        
        # 处理文本中的特殊字符
        text = node.text.replace('"', '\\"')
        
        if depth == 0:
            if node.text == "__root__":
                # 虚拟根节点，直接处理子节点
                for child in node.children:
                    lines.append(self._node_to_mermaid(child, 0))
            else:
                lines.append(f'{indent}root(("{text}"))')
                for child in node.children:
                    lines.append(self._node_to_mermaid(child, depth + 1))
        else:
            # 根据深度使用不同的形状
            if depth == 1:
                lines.append(f'{indent}("{text}")')
            elif depth == 2:
                lines.append(f'{indent}["{text}"]')
            else:
                lines.append(f'{indent}){text}(')
            
            for child in node.children:
                lines.append(self._node_to_mermaid(child, depth + 1))
        
        return '\n'.join(lines)
    
    def _generate_mermaid(self, text: str) -> str:
        """生成 Mermaid mindmap 代码"""
        root = self._parse_list(text)
        if not root:
            return ""
        
        mermaid_code = "mindmap\n"
        mermaid_code += self._node_to_mermaid(root)
        
        return mermaid_code
    
    def _preview(self):
        """预览转换结果"""
        input_text = self.input_text.get("1.0", "end-1c")
        if not input_text.strip():
            messagebox.showwarning("警告", "请输入 Markdown 列表内容")
            return
        
        mermaid_code = self._generate_mermaid(input_text)
        
        self.output_text.delete("1.0", END)
        self.output_text.insert("1.0", mermaid_code)
    
    def _insert_mermaid(self):
        """插入 Mermaid 代码到文档"""
        input_text = self.input_text.get("1.0", "end-1c")
        if not input_text.strip():
            messagebox.showwarning("警告", "请输入 Markdown 列表内容")
            return
        
        mermaid_code = self._generate_mermaid(input_text)
        if not mermaid_code:
            messagebox.showerror("错误", "转换失败，请检查列表格式")
            return
        
        # 插入为代码块
        insert_code = f"\n```mermaid\n{mermaid_code}\n```\n"
        
        textbox = self._get_textbox()
        if textbox:
            textbox.insert("insert", insert_code)
            self.dialog.destroy()
            self.dialog = None
            messagebox.showinfo("成功", "思维导图代码已插入")
    
    def _copy_code(self):
        """复制代码到剪贴板"""
        code = self.output_text.get("1.0", "end-1c")
        if not code.strip():
            self._preview()
            code = self.output_text.get("1.0", "end-1c")
        
        if code.strip():
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(code)
            messagebox.showinfo("成功", "代码已复制到剪贴板")
    
    def _load_from_selection(self):
        """从编辑器选中文本加载"""
        textbox = self._get_textbox()
        if textbox:
            try:
                selected = textbox.get("sel.first", "sel.last")
                if selected:
                    self.input_text.delete("1.0", END)
                    self.input_text.insert("1.0", selected)
                    self._preview()
            except Exception:
                messagebox.showinfo("提示", "请先在编辑器中选中列表文本")
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
