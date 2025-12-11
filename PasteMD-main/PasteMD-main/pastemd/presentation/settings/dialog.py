"""Settings configuration dialog."""

import copy
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Callable, Dict, Any

from ...utils.logging import log
from ...utils.win32 import get_dpi_scale
from ...utils.resources import resource_path
from ...i18n import t, iter_languages, get_language_label
from ...core.state import app_state
from ...config.loader import ConfigLoader
from ...config.defaults import DEFAULT_CONFIG


class SettingsDialog:
    """设置对话框"""
    
    def __init__(self, on_save: Callable[[], None], on_close: Optional[Callable[[], None]] = None):
        """
        初始化设置对话框
        
        Args:
            on_save: 保存回调函数
            on_close: 关闭回调函数
        """
        self.on_save_callback = on_save
        self.on_close_callback = on_close
        self.config_loader = ConfigLoader()
        
        # 加载当前配置的副本，避免直接修改 app_state
        self.current_config = copy.deepcopy(app_state.config)
        
        if app_state.root:
            self.root = tk.Toplevel(app_state.root)
        else:
            self.root = tk.Tk()
            
        self.root.title(t("settings.dialog.title"))
        # 默认不强制置顶，便于最小化
        self.root.attributes("-topmost", False)
        # 确保有最小化按钮
        self.root.resizable(True, True)
        self.root.attributes("-toolwindow", False)
        # 设置图标
        try:
            icon_path = resource_path("assets/icons/logo.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            log(f"Failed to set settings dialog icon: {e}")
        
        # 适配高分屏
        scale = get_dpi_scale()
        width = int(600 * scale)
        height = int(500 * scale)
        self.root.geometry(f"{width}x{height}")
        
        # 设置关闭窗口时的处理
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 窗口居中
        self._center_window()
        
        # 创建UI组件
        self._create_widgets()
    
    def is_alive(self) -> bool:
        """判断窗口是否仍存在"""
        try:
            return bool(self.root.winfo_exists())
        except Exception:
            return False
    
    def restore_and_focus(self):
        """从最小化恢复并置于前台"""
        if not self.is_alive():
            return
        
        try:
            self.root.deiconify()
            original_topmost = bool(self.root.attributes("-topmost"))
            # 通过临时取消/恢复置顶让窗口浮到最前
            self.root.attributes("-topmost", False)
            self.root.lift()
            self.root.focus_force()
            self.root.attributes("-topmost", original_topmost)
        except Exception as e:
            log(f"Failed to restore settings dialog: {e}")
        
    def _center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def _create_widgets(self):
        """创建UI组件"""
        # 创建 Notebook (选项卡容器)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建选项卡
        self._create_general_tab()
        self._create_conversion_tab()
        self._create_advanced_tab()
        self._create_experimental_tab()
        
        # 避免首次打开时就选中首个输入框
        self.root.after_idle(self._clear_initial_selection)
        
        # 底部按钮栏
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        cancel_btn = ttk.Button(
            button_frame,
            text=t("settings.buttons.cancel"),
            command=self._on_cancel,
            width=10
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        save_btn = ttk.Button(
            button_frame,
            text=t("settings.buttons.save"),
            command=self._on_save,
            width=10
        )
        save_btn.pack(side=tk.RIGHT, padx=5)

    def _create_general_tab(self):
        """创建常规设置选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(1, weight=1)
        self.notebook.add(frame, text=t("settings.tab.general"))
        
        # 保存目录
        ttk.Label(frame, text=t("settings.general.save_dir")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # 文件路径输入框
        self.save_dir_var = tk.StringVar(value=self.current_config.get("save_dir", ""))
        self.save_dir_entry = ttk.Entry(frame, textvariable=self.save_dir_var, width=50)
        self.save_dir_entry.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        self.save_dir_entry.bind("<FocusIn>", self._on_focus_in)
        
        # 多按钮放在输入框下方
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(button_frame, text=t("settings.general.browse"), command=self._browse_save_dir).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text=t("settings.general.restore_default"), command=self._restore_default_save_dir).pack(side=tk.LEFT, padx=2)
        
        # 复选框选项
        self.keep_file_var = tk.BooleanVar(value=self.current_config.get("keep_file", False))
        ttk.Checkbutton(frame, text=t("settings.general.keep_file"), variable=self.keep_file_var).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.notify_var = tk.BooleanVar(value=self.current_config.get("notify", True))
        ttk.Checkbutton(frame, text=t("settings.general.notify"), variable=self.notify_var).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.auto_open_var = tk.BooleanVar(value=self.current_config.get("auto_open_on_no_app", True))
        ttk.Checkbutton(frame, text=t("settings.general.auto_open"), variable=self.auto_open_var).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.move_cursor_var = tk.BooleanVar(value=self.current_config.get("move_cursor_to_end", True))
        ttk.Checkbutton(frame, text=t("settings.general.move_cursor"), variable=self.move_cursor_var).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 语言设置（移动到常规页最下方）
        ttk.Label(frame, text=t("settings.general.language")).grid(row=6, column=0, sticky=tk.W, pady=(15, 5))
        
        # 获取当前语言代码和对应的显示名称
        current_code = self.current_config.get("language", "zh")
        current_label = get_language_label(current_code)
        
        self.lang_var = tk.StringVar(value=current_label)
        self.lang_combo = ttk.Combobox(frame, textvariable=self.lang_var, state="readonly")
        
        # 构建语言列表
        langs = []
        self.lang_map: Dict[str, str] = {}
        for code, label in iter_languages():
            langs.append(label)
            self.lang_map[label] = code
        
        self.lang_combo['values'] = langs
        self.lang_combo.grid(row=6, column=1, sticky=tk.W, padx=5, pady=(15, 5))
        # 绑定 FocusIn，避免自动全选
        self.lang_combo.bind("<FocusIn>", self._on_focus_in)

    def _create_conversion_tab(self):
        """创建转换设置选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(1, weight=1)
        self.notebook.add(frame, text=t("settings.tab.conversion"))
        
        # Pandoc 路径
        ttk.Label(frame, text=t("settings.conversion.pandoc_path")).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.pandoc_path_var = tk.StringVar(value=self.current_config.get("pandoc_path", "pandoc"))
        self.pandoc_entry = ttk.Entry(frame, textvariable=self.pandoc_path_var, width=50)
        self.pandoc_entry.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        self.pandoc_entry.bind("<FocusIn>", self._on_focus_in)
        
        # 多按钮放在输入框下方
        pandoc_button_frame = ttk.Frame(frame)
        pandoc_button_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(pandoc_button_frame, text=t("settings.general.browse"), command=self._browse_pandoc).pack(side=tk.LEFT, padx=2)
        ttk.Button(pandoc_button_frame, text=t("settings.general.restore_default"), command=self._restore_default_pandoc_path).pack(side=tk.LEFT, padx=2)
        
        # Reference Docx
        ttk.Label(frame, text=t("settings.conversion.reference_docx")).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        ref_docx = self.current_config.get("reference_docx")
        self.ref_docx_var = tk.StringVar(value=ref_docx if ref_docx else "")
        self.ref_docx_entry = ttk.Entry(frame, textvariable=self.ref_docx_var, width=50)
        self.ref_docx_entry.grid(row=2, column=1, sticky=tk.EW, pady=5, padx=5)
        self.ref_docx_entry.bind("<FocusIn>", self._on_focus_in)
        
        ref_button_frame = ttk.Frame(frame)
        ref_button_frame.grid(row=3, column=1, sticky=tk.W, padx=5, pady=(0, 8))
        ttk.Button(ref_button_frame, text=t("settings.general.browse"), command=self._browse_ref_docx).pack(side=tk.LEFT, padx=2)
        ttk.Button(ref_button_frame, text=t("settings.general.clear"), command=self._clear_ref_docx).pack(side=tk.LEFT, padx=2)
        
        # HTML 格式化
        ttk.Label(frame, text=t("settings.conversion.html_formatting"), font=("", 10, "bold")).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(15, 5))
        
        html_fmt = self.current_config.get("html_formatting", {})
        self.strikethrough_var = tk.BooleanVar(value=html_fmt.get("strikethrough_to_del", True))
        ttk.Checkbutton(frame, text=t("settings.conversion.strikethrough"), variable=self.strikethrough_var).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        ttk.Label(frame, text=t("settings.conversion.first_paragraph_heading"), font=("", 10, "bold")).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(12, 5))
        
        # 其他转换选项
        self.md_indent_var = tk.BooleanVar(value=self.current_config.get("md_disable_first_para_indent", True))
        ttk.Checkbutton(frame, text=t("settings.conversion.md_indent"), variable=self.md_indent_var).grid(row=7, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        self.html_indent_var = tk.BooleanVar(value=self.current_config.get("html_disable_first_para_indent", True))
        ttk.Checkbutton(frame, text=t("settings.conversion.html_indent"), variable=self.html_indent_var).grid(row=8, column=0, columnspan=3, sticky=tk.W, pady=2)

    def _create_advanced_tab(self):
        """创建高级设置选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text=t("settings.tab.advanced"))
        
        # Excel 选项
        self.excel_enable_var = tk.BooleanVar(value=self.current_config.get("enable_excel", True))
        ttk.Checkbutton(frame, text=t("settings.advanced.excel_enable"), variable=self.excel_enable_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.excel_format_var = tk.BooleanVar(value=self.current_config.get("excel_keep_format", True))
        ttk.Checkbutton(frame, text=t("settings.advanced.excel_format"), variable=self.excel_format_var).grid(row=1, column=0, sticky=tk.W, pady=5)

    def _create_experimental_tab(self):
        """创建实验性功能选项卡"""
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text=t("settings.tab.experimental"))
        
        self.keep_formula_var = tk.BooleanVar(value=self.current_config.get("Keep_original_formula", False))
        ttk.Checkbutton(frame, text=t("settings.conversion.keep_formula"), variable=self.keep_formula_var).grid(row=0, column=0, sticky=tk.W, pady=5)

    def _browse_save_dir(self):
        """浏览保存目录"""
        path = filedialog.askdirectory(
            title=t("settings.dialog.select_save_dir"),
            initialdir=os.path.expandvars(self.save_dir_var.get())
        )
        
        if path:
            self.save_dir_var.set(path)

    def _restore_default_save_dir(self):
        """恢复默认保存目录（展开环境变量）"""
        default_save_dir = DEFAULT_CONFIG["save_dir"]
        # 展开环境变量，如 %USERPROFILE% 转换为实际路径
        expanded_save_dir = os.path.expandvars(default_save_dir)
        self.save_dir_var.set(expanded_save_dir)
    
    def _restore_default_pandoc_path(self):
        """恢复默认 Pandoc 路径"""
        default_path = DEFAULT_CONFIG.get("pandoc_path", "pandoc")
        self.pandoc_path_var.set(default_path)
    
    def _clear_ref_docx(self):
        """清除参考文档路径"""
        self.ref_docx_var.set("")
    
    def _clear_initial_selection(self):
        """打开窗口时避免一开始就选中首个输入框"""
        try:
            self.notebook.focus_set()
        except Exception as e:
            log(f"Failed to clear initial selection: {e}")

    def _browse_pandoc(self):
        """浏览 Pandoc 可执行文件"""
        path = filedialog.askopenfilename(
            title=t("settings.dialog.select_pandoc"),
            filetypes=[(t("settings.file_type.executable"), "*.exe"), (t("settings.file_type.all_files"), "*.*")],
            initialdir=os.path.dirname(self.pandoc_path_var.get()) if self.pandoc_path_var.get() else None
        )
        
        if path:
            self.pandoc_path_var.set(path)

    def _browse_ref_docx(self):
        """浏览参考 docx 文件"""
        path = filedialog.askopenfilename(
            title=t("settings.dialog.select_ref_docx"),
            filetypes=[(t("settings.file_type.word_doc"), "*.docx"), (t("settings.file_type.all_files"), "*.*")],
            initialdir=os.path.dirname(self.ref_docx_var.get()) if self.ref_docx_var.get() else None
        )
        
        if path:
            self.ref_docx_var.set(path)

    def _on_save(self):
        """保存配置"""
        try:
            # 更新配置字典
            new_config = self.current_config.copy()
            
            # 将显示名称映射回代码
            selected_label = self.lang_var.get()
            new_config["language"] = self.lang_map.get(selected_label, "zh")
            new_config["save_dir"] = self.save_dir_var.get()
            new_config["keep_file"] = self.keep_file_var.get()
            new_config["notify"] = self.notify_var.get()
            new_config["auto_open_on_no_app"] = self.auto_open_var.get()
            new_config["move_cursor_to_end"] = self.move_cursor_var.get()
            
            new_config["pandoc_path"] = self.pandoc_path_var.get()
            ref_docx = self.ref_docx_var.get()
            new_config["reference_docx"] = ref_docx if ref_docx else None
            
            if "html_formatting" not in new_config:
                new_config["html_formatting"] = {}
            new_config["html_formatting"]["strikethrough_to_del"] = self.strikethrough_var.get()
            
            new_config["md_disable_first_para_indent"] = self.md_indent_var.get()
            new_config["html_disable_first_para_indent"] = self.html_indent_var.get()
            new_config["Keep_original_formula"] = self.keep_formula_var.get()
            
            new_config["enable_excel"] = self.excel_enable_var.get()
            new_config["excel_keep_format"] = self.excel_format_var.get()
            
            # 保存到文件
            self.config_loader.save(new_config)
            
            # 更新全局状态
            app_state.config = new_config
            
            # 显示成功消息（置顶）
            self._show_topmost_message(t("settings.title.success"), t("settings.success.saved"), "info")
            
            if self.on_save_callback:
                self.on_save_callback()
                
            self._safe_destroy()
            
        except Exception as e:
            log(f"Failed to save settings: {e}")
            # 显示错误消息（置顶）
            self._show_topmost_message(t("settings.title.error"), t("settings.error.save_failed", error=str(e)), "error")

    def _on_cancel(self):
        self._safe_destroy()

    def _on_close(self):
        if self.on_close_callback:
            self.on_close_callback()
        self._safe_destroy()

    def _safe_destroy(self):
        try:
            self.root.destroy()
        except Exception as e:
            log(f"Error destroying settings window: {e}")

    def _show_topmost_message(self, title, message, msg_type="info"):
        """显示置顶消息框（使用标准样式）"""
        # 创建临时窗口来控制消息框的置顶属性
        temp_root = tk.Toplevel(self.root)
        temp_root.withdraw()  # 隐藏临时窗口
        temp_root.attributes("-topmost", True)
        
        # 使用标准消息框
        if msg_type == "error":
            messagebox.showerror(title, message, parent=self.root)
        else:
            messagebox.showinfo(title, message, parent=self.root)
        
        # 清理临时窗口
        temp_root.destroy()

    def show(self):
        """显示设置对话框（非阻塞模式）"""
        try:
            # 设置对话框在后台显示，让主事件循环继续运行（允许最小化）
            if app_state.root and app_state.root.winfo_exists():
                self.root.transient(app_state.root)
            else:
                self.root.transient(None)
            
            # 监听退出事件，在后台检查
            self._monitor_quit_event()
            
            # 显示窗口
            self.restore_and_focus()
            
        except Exception as e:
            log(f"Error in settings show: {e}")
            self._safe_destroy()
    
    def _monitor_quit_event(self):
        """监听退出事件，自动关闭设置对话框"""
        def check_quit():
            try:
                # 检查退出事件
                if app_state.quit_event and app_state.quit_event.is_set():
                    self._safe_destroy()
                    return
                
                # 继续监听
                self.root.after(200, check_quit)
            except Exception as e:
                log(f"Error in quit monitor: {e}")
                self.root.after(200, check_quit)
                
        if self.is_alive():
            # 开始监听
            self.root.after(200, check_quit)
    
    def _on_focus_in(self, event):
        """所有 Entry/Combobox 获得焦点时，自动清除选中并将光标移到末尾"""
        widget = event.widget

        def clear_sel():
            # 清除自动选中的部分
            try:
                if hasattr(widget, "select_clear"):
                    widget.select_clear()
                elif hasattr(widget, "selection_clear"):
                    widget.selection_clear(0, "end")
            except Exception:
                pass

            # 光标移到末尾
            try:
                if hasattr(widget, "icursor"):
                    widget.icursor("end")
            except Exception:
                pass

        # 立即清一次
        clear_sel()
        # 再用 after_idle 清一次，覆盖类绑定可能重新设置的选中
        try:
            self.root.after_idle(clear_sel)
        except Exception:
            pass
