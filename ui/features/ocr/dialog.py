# -*- coding: utf-8 -*-
"""OCR 对话框界面模块"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, List
import threading

try:
    import customtkinter as ctk
except ImportError:
    ctk = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

from .image_input import ImageInputManager, ImageLoadError
from .ocr_engine import OCREngine, OCRResult, EngineInitError, RecognitionError
from .markdown_gen import MarkdownGenerator
from .batch_processor import BatchOCRProcessor, BatchProgress


class OCRDialog:
    """OCR 功能对话框"""
    
    def __init__(self, app):
        self.app = app
        self.image_input = ImageInputManager(app)
        self.ocr_engine = OCREngine()
        self.markdown_gen = MarkdownGenerator()
        self.batch_processor = BatchOCRProcessor(self.ocr_engine)
        
        self.current_image: Optional['Image.Image'] = None
        self.current_result: Optional[OCRResult] = None
        self.dialog: Optional[ctk.CTkToplevel] = None
        self._processing = False
    
    def show(self) -> None:
        """显示 OCR 对话框"""
        if ctk is None:
            messagebox.showerror("错误", "CustomTkinter 未安装")
            return
        
        if self.dialog is not None and self.dialog.winfo_exists():
            self.dialog.focus()
            return
        
        self._create_ui()
    
    def _create_ui(self) -> None:
        """创建界面"""
        self.dialog = ctk.CTkToplevel(self.app)
        self.dialog.title("📷 OCR 图片转 Markdown")
        self.dialog.geometry("1000x700")
        self.dialog.minsize(800, 600)
        
        # 主容器
        main_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 顶部工具栏
        self._create_toolbar(main_frame)
        
        # 中间内容区
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, pady=(10, 0))
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_rowconfigure(0, weight=1)
        
        # 左侧：图片预览
        self._create_image_panel(content_frame)
        
        # 右侧：结果编辑
        self._create_result_panel(content_frame)
        
        # 底部操作栏
        self._create_action_bar(main_frame)
        
        # 状态栏
        self._create_status_bar(main_frame)
    
    def _create_toolbar(self, parent) -> None:
        """创建工具栏"""
        toolbar = ctk.CTkFrame(parent, fg_color="transparent", height=40)
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        
        # 导入按钮组
        import_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        import_frame.pack(side="left")
        
        self.file_btn = ctk.CTkButton(
            import_frame,
            text="📂 打开文件",
            command=self._on_file_select,
            width=100
        )
        self.file_btn.pack(side="left", padx=(0, 5))
        
        self.clipboard_btn = ctk.CTkButton(
            import_frame,
            text="📋 粘贴",
            command=self._on_clipboard_paste,
            width=80
        )
        self.clipboard_btn.pack(side="left", padx=5)
        
        self.batch_btn = ctk.CTkButton(
            import_frame,
            text="📁 批量导入",
            command=self._on_batch_import,
            width=100
        )
        self.batch_btn.pack(side="left", padx=5)
        
        # 识别选项
        options_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        options_frame.pack(side="right")
        
        self.detect_table_var = ctk.BooleanVar(value=True)
        self.table_check = ctk.CTkCheckBox(
            options_frame,
            text="表格识别",
            variable=self.detect_table_var
        )
        self.table_check.pack(side="left", padx=5)
        
        self.detect_formula_var = ctk.BooleanVar(value=True)
        self.formula_check = ctk.CTkCheckBox(
            options_frame,
            text="公式识别",
            variable=self.detect_formula_var
        )
        self.formula_check.pack(side="left", padx=5)
    
    def _create_image_panel(self, parent) -> None:
        """创建图片预览面板"""
        image_frame = ctk.CTkFrame(parent)
        image_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 标题
        title_label = ctk.CTkLabel(
            image_frame,
            text="📷 图片预览",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        title_label.pack(pady=(10, 5))
        
        # 图片显示区域
        self.image_canvas = ctk.CTkCanvas(
            image_frame,
            bg="#2b2b2b",
            highlightthickness=0
        )
        self.image_canvas.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 提示文字
        self.image_canvas.create_text(
            200, 150,
            text="拖拽图片到此处\n或点击上方按钮导入",
            fill="#888888",
            font=("Microsoft YaHei", 12),
            tags="placeholder"
        )
        
        # 绑定拖拽
        self.image_canvas.bind("<Configure>", self._on_canvas_resize)
    
    def _create_result_panel(self, parent) -> None:
        """创建结果编辑面板"""
        result_frame = ctk.CTkFrame(parent)
        result_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        # 标题
        title_label = ctk.CTkLabel(
            result_frame,
            text="📝 识别结果",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        title_label.pack(pady=(10, 5))
        
        # 结果文本框
        self.result_text = ctk.CTkTextbox(
            result_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word"
        )
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 置信度信息
        self.confidence_label = ctk.CTkLabel(
            result_frame,
            text="",
            font=ctk.CTkFont(size=11)
        )
        self.confidence_label.pack(pady=(0, 10))
    
    def _create_action_bar(self, parent) -> None:
        """创建操作栏"""
        action_frame = ctk.CTkFrame(parent, fg_color="transparent", height=50)
        action_frame.pack(fill="x", pady=(10, 0))
        action_frame.pack_propagate(False)
        
        # 左侧：识别按钮
        left_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        left_frame.pack(side="left")
        
        self.recognize_btn = ctk.CTkButton(
            left_frame,
            text="🔍 开始识别",
            command=self._on_recognize,
            width=120,
            fg_color="#2563eb",
            hover_color="#1d4ed8"
        )
        self.recognize_btn.pack(side="left", padx=(0, 10))
        
        # 右侧：导出按钮
        right_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
        right_frame.pack(side="right")
        
        self.copy_btn = ctk.CTkButton(
            right_frame,
            text="📋 复制",
            command=self._on_copy,
            width=80
        )
        self.copy_btn.pack(side="left", padx=5)
        
        self.insert_btn = ctk.CTkButton(
            right_frame,
            text="📥 插入到文档",
            command=self._on_insert,
            width=120,
            fg_color="#16a34a",
            hover_color="#15803d"
        )
        self.insert_btn.pack(side="left", padx=5)
        
        self.export_btn = ctk.CTkButton(
            right_frame,
            text="💾 保存",
            command=self._on_export,
            width=80
        )
        self.export_btn.pack(side="left", padx=(5, 0))
    
    def _create_status_bar(self, parent) -> None:
        """创建状态栏"""
        self.status_frame = ctk.CTkFrame(parent, fg_color="transparent", height=25)
        self.status_frame.pack(fill="x", pady=(5, 0))
        self.status_frame.pack_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="就绪",
            font=ctk.CTkFont(size=11)
        )
        self.status_label.pack(side="left")
        
        # 进度条（默认隐藏）
        self.progress_bar = ctk.CTkProgressBar(self.status_frame, width=200)
        self.progress_bar.set(0)
    
    def _on_file_select(self) -> None:
        """文件选择回调"""
        file_path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.webp *.gif"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self._load_image(file_path)
    
    def _on_clipboard_paste(self) -> None:
        """剪贴板粘贴回调"""
        try:
            image = self.image_input.load_from_clipboard()
            if image:
                self.current_image = image
                self._display_image(image)
                self._update_status("已从剪贴板加载图片")
            else:
                self._update_status("剪贴板中没有图片")
        except ImageLoadError as e:
            self._update_status(f"加载失败: {e}")
    
    def _on_batch_import(self) -> None:
        """批量导入回调"""
        file_paths = filedialog.askopenfilenames(
            title="选择多张图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.webp *.gif"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_paths:
            self._process_batch(list(file_paths))
    
    def _on_recognize(self) -> None:
        """开始识别"""
        if self.current_image is None:
            self._update_status("请先导入图片")
            return
        
        if self._processing:
            return
        
        self._processing = True
        self.recognize_btn.configure(state="disabled", text="识别中...")
        self._update_status("正在识别...")
        
        # 在后台线程中执行识别
        thread = threading.Thread(target=self._do_recognize, daemon=True)
        thread.start()
    
    def _do_recognize(self) -> None:
        """执行识别（后台线程）"""
        try:
            # 初始化引擎
            if not self.ocr_engine.is_initialized:
                self.ocr_engine.initialize()
            
            # 预处理图片
            image = self.image_input.convert_to_rgb(self.current_image)
            image = self.image_input.resize_for_ocr(image)
            
            # 执行识别
            result = self.ocr_engine.recognize(
                image,
                detect_tables=self.detect_table_var.get(),
                detect_formulas=self.detect_formula_var.get()
            )
            
            self.current_result = result
            
            # 生成 Markdown
            markdown = self.markdown_gen.generate(result)
            
            # 更新 UI（在主线程）
            self.dialog.after(0, lambda: self._on_recognize_complete(markdown, result))
            
        except EngineInitError as e:
            err_msg = f"引擎初始化失败: {e}"
            self.dialog.after(0, lambda msg=err_msg: self._on_recognize_error(msg))
        except RecognitionError as e:
            err_msg = f"识别失败: {e}"
            self.dialog.after(0, lambda msg=err_msg: self._on_recognize_error(msg))
        except Exception as e:
            err_msg = f"错误: {e}"
            self.dialog.after(0, lambda msg=err_msg: self._on_recognize_error(msg))
    
    def _on_recognize_complete(self, markdown: str, result: OCRResult) -> None:
        """识别完成回调"""
        self._processing = False
        self.recognize_btn.configure(state="normal", text="🔍 开始识别")
        
        # 显示结果
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", markdown)
        
        # 更新置信度
        avg_conf = result.average_confidence * 100
        self.confidence_label.configure(
            text=f"平均置信度: {avg_conf:.1f}% | 识别区域: {len(result.regions)} | 耗时: {result.processing_time:.2f}s"
        )
        
        self._update_status(f"识别完成，共 {len(result.regions)} 个区域")
    
    def _on_recognize_error(self, error: str) -> None:
        """识别错误回调"""
        self._processing = False
        self.recognize_btn.configure(state="normal", text="🔍 开始识别")
        self._update_status(error)
        messagebox.showerror("识别错误", error)
    
    def _on_insert(self) -> None:
        """插入到文档"""
        markdown = self.result_text.get("1.0", "end-1c")
        if not markdown.strip():
            self._update_status("没有可插入的内容")
            return
        
        try:
            # 插入到编辑器
            if hasattr(self.app, 'input_text'):
                self.app.input_text.insert("insert", markdown)
                self.app.on_text_change(None)
                self._update_status("已插入到文档")
        except Exception as e:
            self._update_status(f"插入失败: {e}")
    
    def _on_copy(self) -> None:
        """复制到剪贴板"""
        markdown = self.result_text.get("1.0", "end-1c")
        if not markdown.strip():
            self._update_status("没有可复制的内容")
            return
        
        try:
            self.dialog.clipboard_clear()
            self.dialog.clipboard_append(markdown)
            self._update_status("已复制到剪贴板")
        except Exception as e:
            self._update_status(f"复制失败: {e}")
    
    def _on_export(self) -> None:
        """保存到文件"""
        markdown = self.result_text.get("1.0", "end-1c")
        if not markdown.strip():
            self._update_status("没有可保存的内容")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存 Markdown",
            defaultextension=".md",
            filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(markdown)
                self._update_status(f"已保存: {os.path.basename(file_path)}")
            except Exception as e:
                self._update_status(f"保存失败: {e}")
    
    def _load_image(self, file_path: str) -> None:
        """加载图片"""
        try:
            image = self.image_input.load_from_file(file_path)
            if image:
                self.current_image = image
                self._display_image(image)
                self._update_status(f"已加载: {os.path.basename(file_path)}")
        except ImageLoadError as e:
            self._update_status(f"加载失败: {e}")
            messagebox.showerror("加载错误", str(e))
    
    def _display_image(self, image: 'Image.Image') -> None:
        """显示图片"""
        if ImageTk is None:
            return
        
        # 清除占位符
        self.image_canvas.delete("placeholder")
        self.image_canvas.delete("image")
        
        # 获取画布大小
        canvas_width = self.image_canvas.winfo_width()
        canvas_height = self.image_canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width = 400
            canvas_height = 300
        
        # 计算缩放比例
        img_width, img_height = image.size
        scale = min(canvas_width / img_width, canvas_height / img_height, 1.0)
        
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        
        # 缩放图片
        display_image = image.copy()
        display_image = display_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # 转换为 PhotoImage
        self._photo_image = ImageTk.PhotoImage(display_image)
        
        # 居中显示
        x = canvas_width // 2
        y = canvas_height // 2
        
        self.image_canvas.create_image(x, y, image=self._photo_image, tags="image")
    
    def _on_canvas_resize(self, event) -> None:
        """画布大小改变时重新显示图片"""
        if self.current_image:
            self._display_image(self.current_image)
    
    def _process_batch(self, file_paths: List[str]) -> None:
        """批量处理图片"""
        if self._processing:
            return
        
        self._processing = True
        self.recognize_btn.configure(state="disabled")
        self.progress_bar.pack(side="right", padx=10)
        self.progress_bar.set(0)
        
        def on_progress(progress: BatchProgress):
            self.dialog.after(0, lambda: self._update_batch_progress(progress))
        
        def on_complete(result):
            self.dialog.after(0, lambda: self._on_batch_complete(result))
        
        # 异步处理
        self.batch_processor.process_batch_async(
            file_paths,
            on_progress=on_progress,
            on_complete=on_complete,
            detect_tables=self.detect_table_var.get(),
            detect_formulas=self.detect_formula_var.get()
        )
    
    def _update_batch_progress(self, progress: BatchProgress) -> None:
        """更新批量处理进度"""
        self.progress_bar.set(progress.percentage / 100)
        self._update_status(f"处理中: {progress.completed}/{progress.total} - {os.path.basename(progress.current_file)}")
    
    def _on_batch_complete(self, result) -> None:
        """批量处理完成"""
        self._processing = False
        self.recognize_btn.configure(state="normal")
        self.progress_bar.pack_forget()
        
        # 合并结果
        markdown = self.markdown_gen.merge_results(result.results)
        
        # 显示结果
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", markdown)
        
        # 更新状态
        status = f"批量处理完成: 成功 {result.success_count}, 失败 {result.error_count}, 耗时 {result.total_time:.2f}s"
        self._update_status(status)
        
        if result.errors:
            error_msg = "\n".join(result.errors[:5])
            if len(result.errors) > 5:
                error_msg += f"\n... 还有 {len(result.errors) - 5} 个错误"
            messagebox.showwarning("部分失败", error_msg)
    
    def _update_status(self, message: str) -> None:
        """更新状态栏"""
        self.status_label.configure(text=message)


class OCRFeature:
    """OCR 功能入口"""
    
    def __init__(self, app):
        self.app = app
        self._dialog: Optional[OCRDialog] = None
    
    def show_dialog(self) -> None:
        """显示 OCR 对话框"""
        if self._dialog is None:
            self._dialog = OCRDialog(self.app)
        self._dialog.show()
