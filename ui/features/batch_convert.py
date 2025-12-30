# -*- coding: utf-8 -*-
"""
批量转换功能 - 支持批量导入和转换多个Markdown文件
"""

import os
import threading
from tkinter import filedialog, messagebox
import customtkinter as ctk
from typing import List, Callable


class BatchConvertFeature:
    """批量转换功能管理器"""
    
    def __init__(self, app):
        self.app = app
        self.batch_files: List[str] = []
        self.batch_dialog = None
        self.is_processing = False
        self.cancel_flag = False
        
    def show_batch_convert_dialog(self):
        """显示批量转换对话框"""
        if self.batch_dialog and self.batch_dialog.winfo_exists():
            self.batch_dialog.focus()
            return
            
        self.batch_dialog = ctk.CTkToplevel(self.app)
        self.batch_dialog.title("📁 批量转换")
        self.batch_dialog.geometry("800x600")
        self.batch_dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            self.batch_dialog,
            text="批量 Markdown 转 Word",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 按钮框架
        btn_frame = ctk.CTkFrame(self.batch_dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        
        # 添加文件按钮
        add_files_btn = ctk.CTkButton(
            btn_frame,
            text="➕ 添加文件",
            command=self._add_files,
            width=120
        )
        add_files_btn.pack(side="left", padx=5)
        
        # 添加文件夹按钮
        add_folder_btn = ctk.CTkButton(
            btn_frame,
            text="📁 添加文件夹",
            command=self._add_folder,
            width=120
        )
        add_folder_btn.pack(side="left", padx=5)
        
        # 清空列表按钮
        clear_btn = ctk.CTkButton(
            btn_frame,
            text="🗑️ 清空",
            command=self._clear_files,
            width=100,
            fg_color="#EF4444",
            hover_color="#DC2626"
        )
        clear_btn.pack(side="left", padx=5)
        
        # 文件列表框架
        list_frame = ctk.CTkFrame(self.batch_dialog)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 文件列表（使用文本框显示）
        self.file_listbox = ctk.CTkTextbox(
            list_frame,
            height=300,
            font=ctk.CTkFont(size=12)
        )
        self.file_listbox.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 选项框架
        options_frame = ctk.CTkFrame(self.batch_dialog, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # 合并选项
        self.merge_var = ctk.BooleanVar(value=False)
        merge_check = ctk.CTkCheckBox(
            options_frame,
            text="合并为单个Word文档",
            variable=self.merge_var,
            font=ctk.CTkFont(size=13)
        )
        merge_check.pack(side="left", padx=10)
        
        # 进度条
        self.progress_label = ctk.CTkLabel(
            self.batch_dialog,
            text="准备就绪",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.pack(pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(self.batch_dialog, width=760)
        self.progress_bar.pack(padx=20, pady=5)
        self.progress_bar.set(0)
        
        # 底部按钮
        bottom_btn_frame = ctk.CTkFrame(self.batch_dialog, fg_color="transparent")
        bottom_btn_frame.pack(fill="x", padx=20, pady=20)
        
        # 开始转换按钮
        self.start_btn = ctk.CTkButton(
            bottom_btn_frame,
            text="🚀 开始批量转换",
            command=self._start_batch_convert,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            width=200,
            fg_color="#10B981",
            hover_color="#059669"
        )
        self.start_btn.pack(side="left", expand=True, padx=5)
        
        # 取消按钮
        self.cancel_btn = ctk.CTkButton(
            bottom_btn_frame,
            text="⛔ 取消",
            command=self._cancel_batch,
            height=40,
            width=120,
            state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=5)
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            bottom_btn_frame,
            text="关闭",
            command=self.batch_dialog.destroy,
            height=40,
            width=120,
            fg_color="#6B7280",
            hover_color="#4B5563"
        )
        close_btn.pack(side="left", padx=5)
        
    def _add_files(self):
        """添加文件"""
        files = filedialog.askopenfilenames(
            title="选择Markdown文件",
            filetypes=[("Markdown文件", "*.md *.markdown"), ("所有文件", "*.*")]
        )
        if files:
            for file in files:
                if file not in self.batch_files:
                    self.batch_files.append(file)
            self._update_file_list()
            
    def _add_folder(self):
        """添加文件夹中的所有Markdown文件"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.endswith(('.md', '.markdown')):
                        full_path = os.path.join(root, file)
                        if full_path not in self.batch_files:
                            self.batch_files.append(full_path)
            self._update_file_list()
            
    def _clear_files(self):
        """清空文件列表"""
        self.batch_files.clear()
        self._update_file_list()
        
    def _update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete("1.0", "end")
        if self.batch_files:
            for i, file in enumerate(self.batch_files, 1):
                self.file_listbox.insert("end", f"{i}. {file}\n")
        else:
            self.file_listbox.insert("end", "暂无文件，请添加要转换的Markdown文件...")
            
    def _start_batch_convert(self):
        """开始批量转换"""
        if not self.batch_files:
            messagebox.showwarning("提示", "请先添加要转换的文件！")
            return
            
        # 选择输出目录
        output_dir = filedialog.askdirectory(title="选择输出目录")
        if not output_dir:
            return
            
        self.is_processing = True
        self.cancel_flag = False
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        
        # 在新线程中执行转换
        thread = threading.Thread(
            target=self._batch_convert_worker,
            args=(output_dir,),
            daemon=True
        )
        thread.start()
        
    def _batch_convert_worker(self, output_dir: str):
        """批量转换工作线程"""
        from converter import MarkdownToWordConverter
        
        total = len(self.batch_files)
        success_count = 0
        failed_files = []
        
        try:
            if self.merge_var.get():
                # 合并模式：所有文件合并为一个Word
                self._update_progress(0, total, "正在合并文档...")
                # TODO: 实现合并逻辑
                output_file = os.path.join(output_dir, "merged_output.docx")
                # 这里需要实现合并转换逻辑
            else:
                # 独立模式：每个文件独立转换
                for i, md_file in enumerate(self.batch_files, 1):
                    if self.cancel_flag:
                        self._update_progress(i, total, "已取消")
                        break
                        
                    filename = os.path.basename(md_file)
                    self._update_progress(i, total, f"正在转换: {filename}")
                    
                    try:
                        # 转换文件
                        output_name = os.path.splitext(filename)[0] + ".docx"
                        output_path = os.path.join(output_dir, output_name)
                        
                        converter = MarkdownToWordConverter(
                            base_dir=os.path.dirname(md_file),
                            style=self.app.config.get('export_style', 'standard'),
                            page_size=self.app.config.get('page_size', 'a4')
                        )
                        
                        converter.convert_file(md_file, output_path)
                        success_count += 1
                        
                    except Exception as e:
                        failed_files.append((filename, str(e)))
                        
            # 完成
            if not self.cancel_flag:
                result_msg = f"批量转换完成！\n成功: {success_count}/{total}"
                if failed_files:
                    result_msg += f"\n失败: {len(failed_files)}\n\n失败文件:\n"
                    for fname, error in failed_files[:5]:  # 只显示前5个
                        result_msg += f"- {fname}: {error}\n"
                        
                self.app.after(0, lambda: messagebox.showinfo("完成", result_msg))
                
        except Exception as e:
            self.app.after(0, lambda: messagebox.showerror("错误", f"批量转换失败: {e}"))
            
        finally:
            self.app.after(0, self._reset_ui)
            
    def _update_progress(self, current: int, total: int, message: str):
        """更新进度"""
        progress = current / total if total > 0 else 0
        self.app.after(0, lambda: self.progress_bar.set(progress))
        self.app.after(0, lambda: self.progress_label.configure(text=message))
        
    def _cancel_batch(self):
        """取消批量转换"""
        self.cancel_flag = True
        self.cancel_btn.configure(state="disabled")
        
    def _reset_ui(self):
        """重置UI状态"""
        self.is_processing = False
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="准备就绪")
