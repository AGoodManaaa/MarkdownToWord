# -*- coding: utf-8 -*-
"""
内嵌图表编辑器功能
支持创建柱状图、折线图、饼图，并插入到文档中
"""

import os
import tempfile
import customtkinter as ctk
from tkinter import messagebox, END, filedialog
from typing import List, Optional, Tuple
from dataclasses import dataclass
from ui.dialog_utils import set_dialog_icon

# 尝试导入 matplotlib，如果失败则提供降级方案
try:
    import matplotlib
    matplotlib.use('Agg')  # 非交互式后端
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False


class ChartEditorFeature:
    """内嵌图表编辑器功能"""
    
    CHART_TYPES = [
        ("bar", "柱状图"),
        ("line", "折线图"),
        ("pie", "饼图"),
        ("scatter", "散点图"),
        ("area", "面积图"),
    ]
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.chart_path: Optional[str] = None
    
    def show_dialog(self):
        """显示图表编辑器对话框"""
        if not HAS_MATPLOTLIB:
            messagebox.showerror("错误", "缺少 matplotlib 库，请先安装：\npip install matplotlib")
            return
        
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📊 图表编辑器")
        self.dialog.geometry("800x650")
        self.dialog.transient(self.app)
        set_dialog_icon(self.dialog)
        
        # 居中显示
        self.dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 800) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 650) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 左侧：数据输入
        left_frame = ctk.CTkFrame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # 图表类型选择
        type_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        type_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(type_frame, text="图表类型:").pack(side="left")
        self.chart_type_var = ctk.StringVar(value="bar")
        for value, label in self.CHART_TYPES:
            ctk.CTkRadioButton(
                type_frame, text=label, variable=self.chart_type_var, value=value,
                command=self._update_preview
            ).pack(side="left", padx=5)
        
        # 标题
        title_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        title_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(title_frame, text="图表标题:").pack(side="left")
        self.title_entry = ctk.CTkEntry(title_frame, width=300)
        self.title_entry.pack(side="left", padx=5, fill="x", expand=True)
        self.title_entry.insert(0, "我的图表")
        
        # 数据输入说明
        ctk.CTkLabel(
            left_frame, 
            text="数据输入 (每行一条数据，格式: 标签,数值)",
            anchor="w"
        ).pack(fill="x")
        
        # 数据输入框
        self.data_text = ctk.CTkTextbox(left_frame, height=200)
        self.data_text.pack(fill="both", expand=True, pady=5)
        
        # 示例数据
        sample_data = """一月,120
二月,150
三月,180
四月,200
五月,170
六月,220"""
        self.data_text.insert("1.0", sample_data)
        
        # 样式选项
        style_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        style_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(style_frame, text="配色:").pack(side="left")
        self.color_var = ctk.StringVar(value="viridis")
        color_options = ["viridis", "plasma", "Set2", "Set3", "Pastel1", "tab10"]
        ctk.CTkOptionMenu(
            style_frame, values=color_options, variable=self.color_var,
            width=100, command=lambda x: self._update_preview()
        ).pack(side="left", padx=5)
        
        self.legend_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            style_frame, text="显示图例", variable=self.legend_var,
            command=self._update_preview
        ).pack(side="left", padx=15)
        
        # 按钮
        btn_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(
            btn_frame, text="🔄 预览", width=80,
            command=self._update_preview
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="📥 插入图表", width=100,
            fg_color=("green", "darkgreen"),
            command=self._insert_chart
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="💾 保存图片", width=100,
            command=self._save_chart
        ).pack(side="left")
        
        # 右侧：预览区
        right_frame = ctk.CTkFrame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        
        ctk.CTkLabel(right_frame, text="预览", font=("", 14, "bold")).pack(pady=5)
        
        self.preview_frame = ctk.CTkFrame(right_frame, fg_color="white")
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 初始预览
        self._update_preview()
    
    def _parse_data(self) -> Tuple[List[str], List[float]]:
        """解析输入数据"""
        labels = []
        values = []
        
        text = self.data_text.get("1.0", "end-1c")
        for line in text.strip().split('\n'):
            line = line.strip()
            if not line:
                continue
            
            parts = line.split(',')
            if len(parts) >= 2:
                label = parts[0].strip()
                try:
                    value = float(parts[1].strip())
                    labels.append(label)
                    values.append(value)
                except ValueError:
                    pass
        
        return labels, values
    
    def _update_preview(self, *args):
        """更新图表预览"""
        if not HAS_MATPLOTLIB:
            return
        
        # 清除旧的画布
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        
        labels, values = self._parse_data()
        if not labels or not values:
            ctk.CTkLabel(
                self.preview_frame, 
                text="请输入有效数据",
                text_color="gray"
            ).pack(expand=True)
            return
        
        try:
            # 创建图表
            fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
            
            chart_type = self.chart_type_var.get()
            title = self.title_entry.get()
            colormap = self.color_var.get()
            
            # 获取颜色
            try:
                cmap = plt.get_cmap(colormap)
                colors = [cmap(i / len(labels)) for i in range(len(labels))]
            except Exception:
                colors = None
            
            if chart_type == "bar":
                ax.bar(labels, values, color=colors)
            elif chart_type == "line":
                ax.plot(labels, values, marker='o', linewidth=2, markersize=8)
                ax.fill_between(range(len(labels)), values, alpha=0.3)
            elif chart_type == "pie":
                ax.pie(values, labels=labels, autopct='%1.1f%%', colors=colors)
            elif chart_type == "scatter":
                ax.scatter(range(len(labels)), values, c=range(len(labels)), cmap=colormap, s=100)
                ax.set_xticks(range(len(labels)))
                ax.set_xticklabels(labels)
            elif chart_type == "area":
                ax.fill_between(range(len(labels)), values, alpha=0.7)
                ax.plot(range(len(labels)), values, linewidth=2)
                ax.set_xticks(range(len(labels)))
                ax.set_xticklabels(labels)
            
            ax.set_title(title, fontsize=14, fontweight='bold')
            
            if chart_type != "pie":
                ax.tick_params(axis='x', rotation=45)
            
            plt.tight_layout()
            
            # 在 Tkinter 中显示
            canvas = FigureCanvasTkAgg(fig, master=self.preview_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            
            # 保存临时图片路径
            self._current_fig = fig
            
            plt.close(fig)
            
        except Exception as e:
            ctk.CTkLabel(
                self.preview_frame, 
                text=f"图表生成失败: {e}",
                text_color="red"
            ).pack(expand=True)
    
    def _insert_chart(self):
        """插入图表到文档"""
        labels, values = self._parse_data()
        if not labels or not values:
            messagebox.showwarning("警告", "请先输入有效数据")
            return
        
        try:
            # 生成图表图片
            chart_path = self._generate_chart_image()
            if not chart_path:
                return
            
            # 插入 Markdown 图片引用
            textbox = self._get_textbox()
            if textbox:
                title = self.title_entry.get() or "图表"
                # 使用相对路径或绝对路径
                md_image = f"\n![{title}]({chart_path})\n"
                textbox.insert("insert", md_image)
                
                self.dialog.destroy()
                self.dialog = None
                messagebox.showinfo("成功", f"图表已插入文档\n图片保存位置: {chart_path}")
        except Exception as e:
            messagebox.showerror("错误", f"插入图表失败: {e}")
    
    def _save_chart(self):
        """保存图表为图片"""
        labels, values = self._parse_data()
        if not labels or not values:
            messagebox.showwarning("警告", "请先输入有效数据")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存图表",
            defaultextension=".png",
            filetypes=[("PNG图片", "*.png"), ("JPEG图片", "*.jpg"), ("SVG矢量图", "*.svg")],
            parent=self.dialog
        )
        if not file_path:
            return
        
        try:
            chart_path = self._generate_chart_image(file_path)
            if chart_path:
                messagebox.showinfo("成功", f"图表已保存: {chart_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def _generate_chart_image(self, output_path: str = None) -> Optional[str]:
        """生成图表图片"""
        labels, values = self._parse_data()
        
        # 创建高质量图表
        fig, ax = plt.subplots(figsize=(8, 6), dpi=150)
        
        chart_type = self.chart_type_var.get()
        title = self.title_entry.get()
        colormap = self.color_var.get()
        
        try:
            cmap = plt.get_cmap(colormap)
            colors = [cmap(i / len(labels)) for i in range(len(labels))]
        except Exception:
            colors = None
        
        if chart_type == "bar":
            ax.bar(labels, values, color=colors)
        elif chart_type == "line":
            ax.plot(labels, values, marker='o', linewidth=2, markersize=8)
            ax.fill_between(range(len(labels)), values, alpha=0.3)
        elif chart_type == "pie":
            ax.pie(values, labels=labels, autopct='%1.1f%%', colors=colors)
        elif chart_type == "scatter":
            ax.scatter(range(len(labels)), values, c=range(len(labels)), cmap=colormap, s=100)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels)
        elif chart_type == "area":
            ax.fill_between(range(len(labels)), values, alpha=0.7)
            ax.plot(range(len(labels)), values, linewidth=2)
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels)
        
        ax.set_title(title, fontsize=16, fontweight='bold')
        
        if chart_type != "pie":
            ax.tick_params(axis='x', rotation=45)
        
        plt.tight_layout()
        
        # 保存图片
        if output_path is None:
            # 使用临时目录或当前文档目录
            if hasattr(self.app, 'current_file') and self.app.current_file:
                base_dir = os.path.dirname(self.app.current_file)
            else:
                base_dir = tempfile.gettempdir()
            
            # 生成唯一文件名
            import time
            filename = f"chart_{int(time.time())}.png"
            output_path = os.path.join(base_dir, filename)
        
        fig.savefig(output_path, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        return output_path
    
    def _get_textbox(self):
        """获取编辑器文本框"""
        try:
            if hasattr(self.app, 'input_text') and self.app.input_text is not None:
                return getattr(self.app.input_text, '_textbox', self.app.input_text)
        except Exception:
            pass
        return None
