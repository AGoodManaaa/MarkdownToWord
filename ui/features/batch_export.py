# -*- coding: utf-8 -*-
"""
批量导出功能
支持选择多个Markdown文件，并行导出为Word/PDF，显示进度和错误报告
"""

import os
import threading
import queue
import customtkinter as ctk
from tkinter import filedialog, messagebox, END
from typing import List, Optional, Callable
from dataclasses import dataclass
from enum import Enum
from concurrent.futures import ThreadPoolExecutor, as_completed
from ui.dialog_utils import set_dialog_icon


class ExportStatus(Enum):
    """导出状态"""
    PENDING = "pending"
    EXPORTING = "exporting"
    SUCCESS = "success"
    FAILED = "failed"


@dataclass
class ExportTask:
    """导出任务"""
    input_path: str
    output_path: str
    format: str  # 'docx' or 'pdf'
    status: ExportStatus = ExportStatus.PENDING
    error_message: str = ""


class BatchExportFeature:
    """批量导出功能"""
    
    def __init__(self, app):
        self.app = app
        self.dialog = None
        self.tasks: List[ExportTask] = []
        self.exporting = False
        self.executor: Optional[ThreadPoolExecutor] = None
    
    def show_dialog(self):
        """显示批量导出对话框"""
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📤 批量导出")
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
            toolbar, text="📁 选择文件", width=100,
            command=self._select_files
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="📂 选择文件夹", width=100,
            command=self._select_folder
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            toolbar, text="🗑️ 清空", width=80,
            command=self._clear_tasks
        ).pack(side="left")
        
        # 导出格式
        format_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        format_frame.pack(side="right")
        
        ctk.CTkLabel(format_frame, text="格式:").pack(side="left", padx=(0, 5))
        self.format_var = ctk.StringVar(value="docx")
        ctk.CTkRadioButton(
            format_frame, text="Word", variable=self.format_var, value="docx"
        ).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(
            format_frame, text="PDF", variable=self.format_var, value="pdf"
        ).pack(side="left")
        
        # 输出目录
        output_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        output_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(output_frame, text="输出目录:").pack(side="left")
        self.output_dir_var = ctk.StringVar(value="与源文件相同目录")
        self.output_entry = ctk.CTkEntry(
            output_frame, textvariable=self.output_dir_var, width=400
        )
        self.output_entry.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(
            output_frame, text="...", width=30,
            command=self._select_output_dir
        ).pack(side="left")
        
        # 进度条
        self.progress_var = ctk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(main_frame, variable=self.progress_var)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)
        
        # 状态标签
        self.status_label = ctk.CTkLabel(
            main_frame, text="请选择要导出的文件", anchor="w"
        )
        self.status_label.pack(fill="x", pady=(0, 10))
        
        # 文件列表
        self.file_list = ctk.CTkTextbox(main_frame, height=250)
        self.file_list.pack(fill="both", expand=True)
        self.file_list.configure(state="disabled")
        
        # 底部按钮
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        self.export_btn = ctk.CTkButton(
            btn_frame, text="🚀 开始导出", width=120,
            fg_color=("green", "darkgreen"),
            command=self._start_export
        )
        self.export_btn.pack(side="left", padx=(0, 10))
        
        self.stop_btn = ctk.CTkButton(
            btn_frame, text="⏹ 停止", width=80,
            fg_color=("red", "darkred"),
            command=self._stop_export,
            state="disabled"
        )
        self.stop_btn.pack(side="left")
        
        # 统计
        self.stats_label = ctk.CTkLabel(btn_frame, text="", anchor="e")
        self.stats_label.pack(side="right")
        
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _select_files(self):
        """选择多个文件"""
        files = filedialog.askopenfilenames(
            title="选择 Markdown 文件",
            filetypes=[("Markdown文件", "*.md *.markdown"), ("所有文件", "*.*")],
            parent=self.dialog
        )
        if files:
            for f in files:
                self._add_task(f)
            self._update_display()
    
    def _select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(
            title="选择包含 Markdown 文件的文件夹",
            parent=self.dialog
        )
        if folder:
            # 递归查找所有 .md 文件
            for root, dirs, files in os.walk(folder):
                for f in files:
                    if f.endswith(('.md', '.markdown')):
                        self._add_task(os.path.join(root, f))
            self._update_display()
    
    def _add_task(self, input_path: str):
        """添加导出任务"""
        # 避免重复
        existing = {t.input_path for t in self.tasks}
        if input_path not in existing:
            output_dir = self.output_dir_var.get()
            if output_dir == "与源文件相同目录":
                output_dir = os.path.dirname(input_path)
            
            basename = os.path.splitext(os.path.basename(input_path))[0]
            fmt = self.format_var.get()
            output_path = os.path.join(output_dir, f"{basename}.{fmt}")
            
            self.tasks.append(ExportTask(
                input_path=input_path,
                output_path=output_path,
                format=fmt
            ))
    
    def _select_output_dir(self):
        """选择输出目录"""
        folder = filedialog.askdirectory(
            title="选择输出目录",
            parent=self.dialog
        )
        if folder:
            self.output_dir_var.set(folder)
            # 更新所有任务的输出路径
            for task in self.tasks:
                basename = os.path.splitext(os.path.basename(task.input_path))[0]
                task.output_path = os.path.join(folder, f"{basename}.{task.format}")
            self._update_display()
    
    def _clear_tasks(self):
        """清空任务列表"""
        self.tasks = []
        self._update_display()
    
    def _update_display(self):
        """更新显示"""
        self.file_list.configure(state="normal")
        self.file_list.delete("1.0", END)
        
        for task in self.tasks:
            # 状态图标
            if task.status == ExportStatus.SUCCESS:
                icon = "✓"
            elif task.status == ExportStatus.FAILED:
                icon = "✗"
            elif task.status == ExportStatus.EXPORTING:
                icon = "⏳"
            else:
                icon = "○"
            
            line = f"{icon} {os.path.basename(task.input_path)}"
            if task.error_message:
                line += f" [{task.error_message}]"
            line += "\n"
            self.file_list.insert(END, line)
        
        self.file_list.configure(state="disabled")
        
        # 更新统计
        total = len(self.tasks)
        success = sum(1 for t in self.tasks if t.status == ExportStatus.SUCCESS)
        failed = sum(1 for t in self.tasks if t.status == ExportStatus.FAILED)
        self.stats_label.configure(text=f"共 {total} 个文件 | 成功 {success} | 失败 {failed}")
        
        # 更新状态
        if total == 0:
            self.status_label.configure(text="请选择要导出的文件")
        else:
            self.status_label.configure(text=f"已添加 {total} 个文件")
    
    def _start_export(self):
        """开始导出"""
        if not self.tasks:
            messagebox.showinfo("提示", "请先选择要导出的文件")
            return
        
        self.exporting = True
        self.export_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        
        # 重置状态
        for task in self.tasks:
            task.status = ExportStatus.PENDING
            task.error_message = ""
        
        # 在线程中执行
        thread = threading.Thread(target=self._export_thread, daemon=True)
        thread.start()
    
    def _export_thread(self):
        """导出线程"""
        total = len(self.tasks)
        completed = 0
        
        for task in self.tasks:
            if not self.exporting:
                break
            
            task.status = ExportStatus.EXPORTING
            self._schedule_update()
            
            try:
                self._export_single(task)
                task.status = ExportStatus.SUCCESS
            except Exception as e:
                task.status = ExportStatus.FAILED
                task.error_message = str(e)[:50]
            
            completed += 1
            progress = completed / total
            self._schedule_progress(progress)
            self._schedule_update()
        
        self.exporting = False
        self._schedule_finish()
    
    def _export_single(self, task: ExportTask):
        """导出单个文件"""
        from converter import MarkdownToWordConverter
        
        # 读取 Markdown 内容
        with open(task.input_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 创建转换器
        base_dir = os.path.dirname(task.input_path)
        converter = MarkdownToWordConverter(base_dir=base_dir)
        doc = converter.convert_text(content)
        
        if task.format == 'docx':
            converter.save(task.output_path)
        elif task.format == 'pdf':
            # 先保存为 docx，然后转换为 pdf
            docx_path = task.output_path.replace('.pdf', '.docx')
            converter.save(docx_path)
            
            # 尝试使用已有的 PDF 导出功能
            try:
                if hasattr(self.app, 'pdf_export_feature') and self.app.pdf_export_feature:
                    self.app.pdf_export_feature.convert_docx_to_pdf(docx_path, task.output_path)
                    # 删除临时 docx
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
                else:
                    raise Exception("PDF导出功能不可用")
            except Exception as e:
                # PDF 转换失败，保留 docx
                task.output_path = docx_path
                raise Exception(f"PDF转换失败，已保存为DOCX: {e}")
    
    def _schedule_update(self):
        """调度UI更新"""
        try:
            self.dialog.after(0, self._update_display)
        except Exception:
            pass
    
    def _schedule_progress(self, value: float):
        """调度进度更新"""
        try:
            self.dialog.after(0, lambda: self.progress_bar.set(value))
        except Exception:
            pass
    
    def _schedule_finish(self):
        """调度完成处理"""
        try:
            self.dialog.after(0, self._on_export_finished)
        except Exception:
            pass
    
    def _on_export_finished(self):
        """导出完成"""
        self.export_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        
        success = sum(1 for t in self.tasks if t.status == ExportStatus.SUCCESS)
        failed = sum(1 for t in self.tasks if t.status == ExportStatus.FAILED)
        
        if failed > 0:
            self.status_label.configure(
                text=f"导出完成：成功 {success}，失败 {failed}",
                text_color=("orange", "yellow")
            )
        else:
            self.status_label.configure(
                text=f"全部导出成功！共 {success} 个文件",
                text_color=("green", "lightgreen")
            )
    
    def _stop_export(self):
        """停止导出"""
        self.exporting = False
        self.status_label.configure(text="已停止导出")
        self.export_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
    
    def _on_close(self):
        """关闭对话框"""
        self.exporting = False
        if self.executor:
            self.executor.shutdown(wait=False)
        self.dialog.destroy()
        self.dialog = None
