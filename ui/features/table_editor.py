# -*- coding: utf-8 -*-
"""
表格功能增强 - 可视化编辑器、Excel导入等
"""

import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
import re


class TableEditorFeature:
    """表格编辑器功能"""
    
    def __init__(self, app):
        self.app = app
        self.table_dialog = None
        self.table_data = []
        self.rows = 3
        self.cols = 3
        self.cells = []
        
    def show_table_editor(self, existing_table: str = None):
        """显示表格编辑器"""
        if self.table_dialog and self.table_dialog.winfo_exists():
            self.table_dialog.focus()
            return
            
        self.table_dialog = ctk.CTkToplevel(self.app)
        self.table_dialog.title("📊 表格编辑器")
        self.table_dialog.geometry("900x650")
        self.table_dialog.transient(self.app)
        
        # 解析现有表格数据（如果有）
        if existing_table:
            self._parse_markdown_table(existing_table)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.table_dialog,
            text="可视化表格编辑器",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=15)
        
        # 工具栏
        toolbar = ctk.CTkFrame(self.table_dialog, fg_color="transparent")
        toolbar.pack(fill="x", padx=20, pady=10)
        
        # 大小设置
        size_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        size_frame.pack(side="left")
        
        ctk.CTkLabel(size_frame, text="行:").pack(side="left", padx=3)
        self.rows_var = ctk.StringVar(value=str(self.rows))
        rows_entry = ctk.CTkEntry(size_frame, width=50, textvariable=self.rows_var)
        rows_entry.pack(side="left", padx=3)
        
        ctk.CTkLabel(size_frame, text="列:").pack(side="left", padx=3)
        self.cols_var = ctk.StringVar(value=str(self.cols))
        cols_entry = ctk.CTkEntry(size_frame, width=50, textvariable=self.cols_var)
        cols_entry.pack(side="left", padx=3)
        
        resize_btn = ctk.CTkButton(
            size_frame,
            text="调整大小",
            command=self._resize_table,
            width=100
        )
        resize_btn.pack(side="left", padx=5)
        
        # 操作按钮
        ctk.CTkButton(
            toolbar,
            text="➕ 添加行",
            command=self._add_row,
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            toolbar,
            text="➖ 删除行",
            command=self._delete_row,
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            toolbar,
            text="➕ 添加列",
            command=self._add_column,
            width=100
        ).pack(side="left", padx=3)
        
        ctk.CTkButton(
            toolbar,
            text="➖ 删除列",
            command=self._delete_column,
            width=100
        ).pack(side="left", padx=3)
        
        # Excel导入
        ctk.CTkButton(
            toolbar,
            text="📥 导入Excel",
            command=self._import_from_excel,
            width=120,
            fg_color="#10B981",
            hover_color="#059669"
        ).pack(side="right", padx=3)
        
        # 表格容器（可滚动）
        table_container = ctk.CTkScrollableFrame(
            self.table_dialog,
            width=850,
            height=400
        )
        table_container.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.table_frame = table_container
        self._create_table_grid()
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(self.table_dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        insert_btn = ctk.CTkButton(
            btn_frame,
            text="✅ 插入到文档",
            command=self._insert_table,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=140,
            height=35
        )
        insert_btn.pack(side="left", padx=5)
        
        preview_btn = ctk.CTkButton(
            btn_frame,
            text="👁️ 预览Markdown",
            command=self._preview_markdown,
            width=160,
            height=35
        )
        preview_btn.pack(side="left", padx=5)
        
        close_btn = ctk.CTkButton(
            btn_frame,
            text="关闭",
            command=self.table_dialog.destroy,
            fg_color="#6B7280",
            hover_color="#4B5563",
            width=100,
            height=35
        )
        close_btn.pack(side="right", padx=5)
        
    def _create_table_grid(self):
        """创建表格网格"""
        # 清空现有内容
        for widget in self.table_frame.winfo_children():
            widget.destroy()
            
        self.cells = []
        
        # 创建单元格
        for i in range(self.rows):
            row_cells = []
            for j in range(self.cols):
                # 第一行使用不同颜色（表头）
                fg_color = "#E0E7FF" if i == 0 else "#F9FAFB"
                
                cell = ctk.CTkEntry(
                    self.table_frame,
                    width=120,
                    height=35,
                    fg_color=fg_color
                )
                cell.grid(row=i, column=j, padx=2, pady=2, sticky="ew")
                
                # 填充现有数据
                if i < len(self.table_data) and j < len(self.table_data[i]):
                    cell.insert(0, self.table_data[i][j])
                elif i == 0:
                    cell.insert(0, f"列{j+1}")
                    
                row_cells.append(cell)
                
            self.cells.append(row_cells)
            
    def _resize_table(self):
        """调整表格大小"""
        try:
            new_rows = int(self.rows_var.get())
            new_cols = int(self.cols_var.get())
            
            if new_rows < 1 or new_cols < 1:
                messagebox.showwarning("提示", "行数和列数必须大于0！")
                return
                
            if new_rows > 50 or new_cols > 20:
                messagebox.showwarning("提示", "表格过大！建议行数≤50，列数≤20")
                return
                
            # 保存当前数据
            self._save_table_data()
            
            self.rows = new_rows
            self.cols = new_cols
            
            # 重新创建表格
            self._create_table_grid()
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")
            
    def _save_table_data(self):
        """保存当前表格数据"""
        self.table_data = []
        for row_cells in self.cells:
            row_data = [cell.get() for cell in row_cells]
            self.table_data.append(row_data)
            
    def _add_row(self):
        """添加行"""
        self._save_table_data()
        self.rows += 1
        self.table_data.append([""] * self.cols)
        self._create_table_grid()
        
    def _delete_row(self):
        """删除最后一行"""
        if self.rows <= 1:
            messagebox.showwarning("提示", "至少需要保留一行（表头）！")
            return
            
        self._save_table_data()
        self.rows -= 1
        if len(self.table_data) >= self.rows:
            self.table_data = self.table_data[:self.rows]
        self._create_table_grid()
        
    def _add_column(self):
        """添加列"""
        self._save_table_data()
        self.cols += 1
        for row in self.table_data:
            row.append("")
        self._create_table_grid()
        
    def _delete_column(self):
        """删除最后一列"""
        if self.cols <= 1:
            messagebox.showwarning("提示", "至少需要保留一列！")
            return
            
        self._save_table_data()
        self.cols -= 1
        for row in self.table_data:
            if len(row) >= self.cols:
                row.pop()
        self._create_table_grid()
        
    def _import_from_excel(self):
        """从Excel导入数据"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            import pandas as pd
            
            # 读取Excel或CSV
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
                
            # 转换为列表
            self.table_data = df.fillna("").astype(str).values.tolist()
            
            # 添加表头
            headers = df.columns.tolist()
            self.table_data.insert(0, [str(h) for h in headers])
            
            self.rows = len(self.table_data)
            self.cols = len(self.table_data[0]) if self.table_data else 0
            
            # 更新UI
            self.rows_var.set(str(self.rows))
            self.cols_var.set(str(self.cols))
            self._create_table_grid()
            
            messagebox.showinfo("成功", f"成功导入 {self.rows-1} 行数据")
            
        except ImportError:
            messagebox.showerror(
                "缺少依赖",
                "需要安装 pandas 和 openpyxl:\n\npip install pandas openpyxl"
            )
        except Exception as e:
            messagebox.showerror("导入失败", f"无法导入文件:\n{e}")
            
    def _generate_markdown_table(self) -> str:
        """生成Markdown表格"""
        self._save_table_data()
        
        if not self.table_data:
            return ""
            
        lines = []
        
        # 表头
        if self.table_data:
            headers = [cell or " " for cell in self.table_data[0]]
            lines.append("| " + " | ".join(headers) + " |")
            
            # 分隔线
            separators = ["---"] * len(headers)
            lines.append("|" + "|".join(separators) + "|")
            
            # 数据行
            for row in self.table_data[1:]:
                cells = [cell or " " for cell in row]
                lines.append("| " + " | ".join(cells) + " |")
                
        return "\n".join(lines)
        
    def _preview_markdown(self):
        """预览Markdown"""
        markdown = self._generate_markdown_table()
        
        # 创建预览对话框
        preview = ctk.CTkToplevel(self.table_dialog)
        preview.title("Markdown 预览")
        preview.geometry("600x400")
        preview.transient(self.table_dialog)
        
        text_box = ctk.CTkTextbox(preview, font=ctk.CTkFont(family="Consolas"))
        text_box.pack(fill="both", expand=True, padx=10, pady=10)
        text_box.insert("1.0", markdown)
        
        copy_btn = ctk.CTkButton(
            preview,
            text="📋 复制",
            command=lambda: self._copy_to_clipboard(markdown),
            width=100
        )
        copy_btn.pack(pady=10)
        
    def _copy_to_clipboard(self, text):
        """复制到剪贴板"""
        self.table_dialog.clipboard_clear()
        self.table_dialog.clipboard_append(text)
        messagebox.showinfo("成功", "已复制到剪贴板")
        
    def _insert_table(self):
        """插入表格到文档"""
        markdown = self._generate_markdown_table()
        
        if markdown:
            self.app.input_text.insert("insert", "\n" + markdown + "\n\n")
            self.app.on_text_change(None)
            messagebox.showinfo("成功", "表格已插入到文档")
            self.table_dialog.destroy()
        else:
            messagebox.showwarning("提示", "表格为空！")
            
    def _parse_markdown_table(self, markdown_table: str):
        """解析Markdown表格"""
        lines = [line.strip() for line in markdown_table.strip().split('\n') if line.strip()]
        
        if len(lines) < 2:
            return
            
        self.table_data = []
        
        for i, line in enumerate(lines):
            # 跳过分隔线
            if i == 1 and re.match(r'^[\|\s\-:]+$', line):
                continue
                
            # 解析单元格
            cells = [cell.strip() for cell in line.split('|')]
            # 移除首尾空元素（由于前后的|产生）
            if cells and cells[0] == '':
                cells = cells[1:]
            if cells and cells[-1] == '':
                cells = cells[:-1]
                
            if cells:
                self.table_data.append(cells)
                
        if self.table_data:
            self.rows = len(self.table_data)
            self.cols = max(len(row) for row in self.table_data)
            
            # 补齐不足的列
            for row in self.table_data:
                while len(row) < self.cols:
                    row.append("")
