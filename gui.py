# -*- coding: utf-8 -*-

import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk


# 导入转换器和主题模块
from ui.theme import COLORS, COLORS_LIGHT, COLORS_DARK, load_config, save_config
from ui.widgets import ModernButton, ModernCard, ExportStyleSettingsDialog
from ui.editor import LineNumberedText
from ui.preview import MarkdownPreview
from ui.dialogs import SearchReplaceDialog
from ui.sidebar import OutlineView, RecentFilesView
from ui.features import (
    TooltipManager,
    HeaderStyler,
    CommandPalette,
    InsertTemplatesFeature,
    StatusBarFeature,
    EditorContextMenuFeature,
    HelpDialogFeature,
    AutoSaveFeature,
    FileOpsFeature,
    ThemeFeature,
    PreviewSyncFeature,
    WindowGeometryFeature,
)
from ui.export_helpers import (
    export_to_word_for_app,
    show_export_options_for_app,
    do_export_for_app,
    on_export_success_for_app,
    on_export_error_for_app,
)
from ui.clipboard import (
    copy_to_clipboard_for_app,
    copy_word_to_clipboard_for_app,
    copy_as_html_for_app,
    copy_markdown_to_clipboard_for_app,
    show_copy_toast_for_app,
)
from ui.export_history import show_export_history_dialog
from ui.busy_state import BusyState
from ui.drag_drop import handle_drop_for_app
from ui.startup_content import insert_example_if_empty_for_app
from ui.formatting import show_format_dialog_for_app


class App(ctk.CTk):
    """主应用窗口 - 优化版"""
    
    def __init__(self):
        super().__init__()
        
        # 窗口配置
        self.title("✨ Markdown → Word 转换器 v2.0")
        self.geometry("1500x900")
        self.minsize(1100, 700)
        
        # 设置窗口图标
        try:
            import os
            icon_path = os.path.join(os.path.dirname(__file__), 'app.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # 图标加载失败不影响程序运行
        
        # 加载配置
        self.config = load_config()

        self.file_ops = FileOpsFeature(self)
        self.theme_feature = ThemeFeature(self)

        self.preview_sync = PreviewSyncFeature(self)
        self.window_geometry_feature = WindowGeometryFeature(self)
        self.theme_feature.apply_mode(self.config.get('theme', 'light'))
        
        # 设置窗口背景
        self.configure(fg_color=COLORS['bg_light'])
        
        # 当前文件路径
        self.current_file = None

        self.busy = BusyState(self)

        # 导出取消标记（线程间通信）
        self._export_cancel_event = None
        
        # 防抖定时器ID
        self._debounce_id = None
        
        # 搜索对话框引用
        self.search_dialog = None

        self._header_default_buttons = []

        # features
        self.tooltip = TooltipManager(self)
        self.header_styler = HeaderStyler(self)
        self.command_palette = CommandPalette(self)

        self.insert_templates = InsertTemplatesFeature(self)
        self.status_bar_feature = StatusBarFeature(self)
        self.editor_context_menu_feature = EditorContextMenuFeature(self)

        self.help_dialog_feature = HelpDialogFeature(self)
        self.auto_save_feature = AutoSaveFeature(self)
        
        # 内容修改标记
        self._content_modified = False
        self._last_saved_content = ""
        self._last_content_snapshot = None
        
        # 构建界面
        self._create_header()
        self._create_status_bar()  # 先创建状态栏
        self._create_main_content()  # 再创建主内容（包含_insert_example调用）

        self.header_styler.update_states()
        
        # 绑定快捷键
        self.bind('<Control-o>', lambda e: self.open_file())
        self.bind('<Control-s>', lambda e: self.save_file())  # 保存源文件
        self.bind('<Control-Shift-s>', lambda e: self.export_to_word())  # 导出Word
        self.bind('<Control-Shift-f>', lambda e: self.format_markdown())
        self.bind('<Control-Shift-c>', lambda e: self.copy_to_clipboard())
        self.bind('<Control-j>', lambda e: self.show_export_history())
        self.bind('<Control-f>', lambda e: self.show_search_dialog())
        self.bind('<Control-h>', lambda e: self.show_search_dialog())
        self.bind('<Control-plus>', lambda e: self.change_font_size(1))
        self.bind('<Control-minus>', lambda e: self.change_font_size(-1))
        self.bind('<Control-b>', lambda e: self.toggle_sidebar())
        self.bind('<Control-p>', lambda e: self.toggle_preview())
        self.bind('<Control-z>', lambda e: self._undo())
        self.bind('<Control-y>', lambda e: self._redo())
        self.bind('<Control-Shift-z>', lambda e: self._redo())
        self.bind('<Control-k>', lambda e: self.command_palette.show())
        self.bind('<F1>', lambda e: self.show_help())
        
        # 绑定窗口关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 支持拖拽导入文件
        self._setup_drag_drop()
        
        # 更新最近文件列表
        self.file_ops.update_recent_files_view()
        
        # 恢复窗口位置和大小
        self._restore_window_geometry()
        
        # 启动自动保存
        self.auto_save_feature.start()
    
    def _create_header(self):
        """创建顶部标题栏"""
        self.header = ctk.CTkFrame(self, fg_color=COLORS['primary'], height=60, corner_radius=0)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)
        
        # 左侧Logo和标题
        left_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        left_frame.pack(side="left", padx=20, pady=12)
        
        self.title_label = ctk.CTkLabel(
            left_frame,
            text="📝 Markdown → Word",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        )
        self.title_label.pack(side="left")
        

        # 中间工具栏
        toolbar_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        toolbar_frame.pack(side="left", padx=24)
        
        # 工具按钮
        tools = [
            ("📂", "打开", self.open_file, "Ctrl+O"),
            ("💾", "保存", self.save_file, "Ctrl+S"),
            ("🧹", "规范化", self.format_markdown, "Ctrl+Shift+F"),
            ("🔍", "搜索", self.show_search_dialog, "Ctrl+F"),
            ("👁", "预览", self.toggle_preview, "Ctrl+P"),
            ("📤", "导出", self.export_to_word, "Ctrl+Shift+S"),
            ("🕘", "历史", self.show_export_history, "Ctrl+J"),
        ]
        
        self.preview_btn = None
        for icon, tip, cmd, shortcut in tools:
            btn = ctk.CTkButton(
                toolbar_frame,
                text=icon,
                width=38,
                height=34,
                corner_radius=10,
                fg_color="transparent",
                hover_color=COLORS['primary_hover'],
                text_color="white",
                font=ctk.CTkFont(size=15, weight="bold"),
                command=cmd,
            )
            btn.pack(side="left", padx=2)
            self._header_default_buttons.append(btn)
            self.tooltip.add_tooltip(btn, f"{tip}\n{shortcut}")
            if tip == "预览":
                self.preview_btn = btn
        
        # 插入按钮（带下拉菜单）
        self.insert_btn = ctk.CTkButton(
            toolbar_frame,
            text="➕",
            width=38,
            height=34,
            corner_radius=10,
            fg_color=COLORS['success'],
            hover_color="#16A34A",
            text_color="white",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda: None  # 占位
        )
        self.insert_btn.pack(side="left", padx=2)
        self.insert_btn.bind('<Button-1>', self.show_insert_menu)
        self.tooltip.add_tooltip(self.insert_btn, "插入\n点击选择模板")
        
        # 右侧按钮组
        btn_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        btn_frame.pack(side="right", padx=20, pady=12)
        
        # 侧边栏切换
        self.sidebar_btn = ctk.CTkButton(
            btn_frame,
            text="☰",
            command=self.toggle_sidebar,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34,
            font=ctk.CTkFont(size=15, weight="bold")
        )
        self.sidebar_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.sidebar_btn)
        self.tooltip.add_tooltip(self.sidebar_btn, "侧边栏\nCtrl+B")
        
        # 字体调整
        self.font_minus_btn = ctk.CTkButton(
            btn_frame,
            text="A-",
            command=lambda: self.change_font_size(-1),
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.font_minus_btn.pack(side="left", padx=1)
        self._header_default_buttons.append(self.font_minus_btn)
        self.tooltip.add_tooltip(self.font_minus_btn, "字体大小\nCtrl++ / Ctrl+-")
        
        self.font_plus_btn = ctk.CTkButton(
            btn_frame,
            text="A+",
            command=lambda: self.change_font_size(1),
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.font_plus_btn.pack(side="left", padx=1)
        self._header_default_buttons.append(self.font_plus_btn)
        self.tooltip.add_tooltip(self.font_plus_btn, "字体大小\nCtrl++ / Ctrl+-")
        
        # 主题切换
        self.theme_btn = ctk.CTkButton(
            btn_frame,
            text=("☀️" if ctk.get_appearance_mode() == "Dark" else "🌙"),
            command=self.toggle_theme,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.theme_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.theme_btn)
        self.tooltip.add_tooltip(self.theme_btn, "主题\n切换明/暗")

        self.export_style_header_btn = ctk.CTkButton(
            btn_frame,
            text="⚙",
            command=self.open_export_style_settings,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34,
        )
        self.export_style_header_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.export_style_header_btn)
        try:
            self.tooltip.add_tooltip(self.export_style_header_btn, "导出样式设置\n(含导入Word模板)")
        except Exception:
            pass

        # 初始样式刷新
        self.header_styler.update_states()
    
    def _create_main_content(self):
        """创建主内容区域 - 包含侧边栏"""
        # 主容器
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 左侧边栏（大纲+最近文件）
        self.sidebar_visible = self.config.get('sidebar_visible', True)
        self.sidebar = ctk.CTkFrame(self.main_container, fg_color=COLORS['bg_sidebar'], width=250, corner_radius=12)
        if self.sidebar_visible:
            self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)
        
        # 侧边栏内容
        self._create_sidebar_content()
        
        # 右侧主编辑区
        self.main_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.main_frame.pack(side="left", fill="both", expand=True)
        
        # 配置列权重：左侧输入略宽，右侧预览略窄
        self.main_frame.grid_columnconfigure(0, weight=3)
        self.main_frame.grid_columnconfigure(1, weight=2)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # ===== 左侧：输入区域 =====
        self._create_input_panel(self.main_frame)
        
        # ===== 右侧：预览区域 =====
        self._create_preview_panel(self.main_frame)
        
        # 插入示例文本（在所有组件创建完成后）
        self._insert_example()
    
    def _create_sidebar_content(self):
        """创建侧边栏内容"""
        # 大纲视图
        self.outline_view = OutlineView(
            self.sidebar,
            on_heading_click=self._jump_to_line
        )
        self.outline_view.pack(fill="both", expand=True, pady=(0, 10))
        
        # 分隔线
        separator = ctk.CTkFrame(self.sidebar, height=1, fg_color=COLORS['border'])
        separator.pack(fill="x", padx=15, pady=5)
        
        # 最近文件
        self.recent_files_view = RecentFilesView(
            self.sidebar,
            on_file_click=self._open_recent_file
        )
        self.recent_files_view.pack(fill="both", expand=True)
    
    def _create_input_panel(self, parent):
        """创建输入面板 - 带行号"""
        self.input_card = ModernCard(parent)
        self.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        
        # 工具栏 - 紧凑布局，紧贴文本框
        toolbar = ctk.CTkFrame(self.input_card, fg_color="transparent", height=26)
        toolbar.pack(fill="x", padx=6, pady=(6, 0))
        toolbar.pack_propagate(False)  # 保持固定高度
        
        # 快捷插入按钮 - 分组显示
        groups = [
            # 标题组
            [("H1", "# "), ("H2", "## "), ("H3", "### ")],
            # 格式组
            [("B", "**粗体**"), ("I", "*斜体*"), ("~", "~~删除~~")],
            # 上下标组
            [("²", "<sup>上标</sup>"), ("₂", "<sub>下标</sub>")],
            # 插入组
            [("🖼", "![图片](url)"), ("🔗", "[链接](url)"), ("∑", "$公式$")],
            # 块级组
            [("≣", "| 表头 |\n|---|\n| 内容 |"), ("`", "```python\ncode\n```")],
        ]
        
        for i, group in enumerate(groups):
            if i > 0:
                # 分隔线
                sep = ctk.CTkFrame(toolbar, width=1, fg_color=COLORS['border'])
                sep.pack(side="left", fill="y", padx=3, pady=2)
            
            for text, insert_text in group:
                btn = ctk.CTkButton(
                    toolbar,
                    text=text,
                    width=26,
                    height=22,
                    corner_radius=8,
                    fg_color=COLORS['bg_card'],
                    text_color=COLORS['text_primary'],
                    hover_color=COLORS['highlight'],
                    border_width=1,
                    border_color=COLORS['border'],
                    font=ctk.CTkFont(size=10, weight="bold"),
                    command=lambda t=insert_text: self.insert_text(t)
                )
                btn.pack(side="left", padx=1)
        
        # 带行号的输入文本框
        self.input_editor = LineNumberedText(
            self.input_card,
            font_size=self.config.get('font_size', 14),
            on_scroll=self._on_editor_scroll  # 滚动同步回调
        )
        self.input_editor.pack(fill="both", expand=True, padx=6, pady=(4, 6))
        
        # 兼容旧属性名
        self.input_text = self.input_editor
        
        # 绑定实时预览（带防抖）
        self.input_editor.bind('<KeyRelease>', self._on_text_change_debounced)
        # Tab 跳转到下一处占位文本；Esc 取消占位跳转
        self.input_editor.bind('<Tab>', self.insert_templates.on_tab)
        self.input_editor.bind('<Escape>', self.insert_templates.on_escape)
        # 光标/选择变化时更新状态栏行列
        self.input_editor.bind('<KeyRelease>', self._on_cursor_event)
        self.input_editor.bind('<ButtonRelease-1>', self._on_cursor_event)

        # 编辑器右键菜单
        try:
            self.editor_context_menu_feature.attach(self.input_editor._textbox)
        except Exception:
            pass
    
    def _create_preview_panel(self, parent):
        """创建预览面板 - 支持开关"""
        self.preview_visible = True
        self.preview_card = ModernCard(parent, title="👁️ 实时预览")
        self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        
        # 预览组件
        self.preview = MarkdownPreview(self.preview_card, on_content_change=self._on_preview_change, app=self)
        self.preview.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        
        # 底部操作按钮
        btn_frame = ctk.CTkFrame(self.preview_card, fg_color="transparent", height=45)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        left_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
        left_group.pack(side="left")

        right_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
        right_group.pack(side="right")
        
        # 导出Word按钮
        self.export_btn = ModernButton(
            left_group,
            text="📤 导出",
            command=self.export_to_word,
            style="primary",
            width=92
        )
        self.export_btn.pack(side="left", padx=(0, 8))

        # 取消导出按钮（导出进行中可用）
        self.cancel_export_btn = ModernButton(
            left_group,
            text="⛔ 取消",
            command=self.cancel_export,
            style="ghost",
            width=86,
        )
        self.cancel_export_btn.pack(side="left", padx=(0, 8))
        try:
            self.cancel_export_btn.configure(state="disabled")
        except Exception:
            pass
        try:
            self.tooltip.add_tooltip(self.cancel_export_btn, "取消导出\n导出进行中可用")
        except Exception:
            pass

        self.export_style_btn = ModernButton(
            left_group,
            text="⚙",
            command=self.open_export_style_settings,
            style="ghost",
            width=36,
        )
        self.export_style_btn.pack(side="left", padx=(0, 6))
        try:
            self.tooltip.add_tooltip(self.export_style_btn, "导出样式设置")
        except Exception:
            pass
        
        self.export_history_btn = ModernButton(
            left_group,
            text="🕘",
            command=self.show_export_history,
            style="ghost",
            width=36,
        )
        self.export_history_btn.pack(side="left", padx=(0, 6))
        try:
            self.tooltip.add_tooltip(self.export_history_btn, "导出历史")
        except Exception:
            pass
        
        # 复制到剪贴板按钮
        self.copy_btn = ModernButton(
            left_group,
            text="📋 复制",
            command=self.copy_to_clipboard,
            style="ghost",
            width=86
        )
        self.copy_btn.pack(side="left", padx=(0, 6))
        
        # 清空按钮
        self.clear_btn = ModernButton(
            right_group,
            text="🗑️",
            command=self.clear_all,
            style="danger",
            width=36
        )
        self.clear_btn.pack(side="right", padx=(0, 8))
    
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_bar_feature.create()
        self.status_bar = self.status_bar_feature.frame
        self.status_label = self.status_bar_feature.status_label
        self.word_count_label = self.status_bar_feature.word_count_label
        self.cursor_pos_label = self.status_bar_feature.cursor_pos_label
    
    def _insert_example(self):
        """插入示例Markdown"""
        insert_example_if_empty_for_app(self)
    
    def insert_text(self, text: str):
        """在光标位置插入文本"""
        self.input_text.insert("insert", text)
        self.on_text_change(None)
    
    def _on_text_change_debounced(self, event):
        self.preview_sync.on_text_change_debounced(event)
    
    def on_text_change(self, event):
        self.preview_sync.on_text_change(event)
    
    def _on_preview_change(self, markdown_text: str):
        self.preview_sync.on_preview_change(markdown_text)

    def open_file(self):
        self.file_ops.open_file()

    def export_to_word(self):
        """导出为Word文档（委托给导出 helper）。"""
        export_to_word_for_app(self)

    def format_markdown(self):
        try:
            show_format_dialog_for_app(self)
        except Exception:
            pass

    def show_export_history(self):
        """显示导出历史。"""
        try:
            show_export_history_dialog(self)
        except Exception:
            pass

    def cancel_export(self):
        """请求取消导出（供导出线程轮询）。"""
        try:
            if self._export_cancel_event is not None:
                self._export_cancel_event.set()
                self.update_status("⛔ 已请求取消导出...")
        except Exception:
            pass

    def _show_export_options(self, content: str):
        """显示导出选项对话框（委托给导出 helper）。"""
        show_export_options_for_app(self, content)

    def _do_export(self, content: str, style: str, page_size: str):
        """执行导出（委托给导出 helper）。"""
        do_export_for_app(self, content, style, page_size)

    def on_export_success(self, file_path):
        """导出成功回调（委托给导出 helper）。"""
        on_export_success_for_app(self, file_path)

    def _open_file_cross_platform(self, file_path: str):
        """跨平台打开文件"""
        import subprocess
        import platform
        
        system = platform.system()
        try:
            if system == 'Windows':
                os.startfile(file_path)
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', file_path], check=True)
            else:  # Linux
                subprocess.run(['xdg-open', file_path], check=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def on_export_error(self, error):
        """导出失败回调（委托给导出 helper）。"""
        on_export_error_for_app(self, error)
    
    def copy_to_clipboard(self):
        """复制内容到剪贴板（委托给剪贴板 helper）。"""
        copy_to_clipboard_for_app(self)

    def copy_markdown_to_clipboard(self, markdown_text: str):
        """复制指定 Markdown 到剪贴板（Word 兼容）。"""
        copy_markdown_to_clipboard_for_app(self, markdown_text)
    
    def _copy_word_to_clipboard(self, docx_path: str):
        """使用 COM 将 Word 内容复制到剪贴板（委托给剪贴板 helper）。"""
        copy_word_to_clipboard_for_app(self, docx_path)
    
    def _copy_as_html(self, docx_path: str):
        """备用方案：转换为 HTML 复制（委托给剪贴板 helper）。"""
        copy_as_html_for_app(self, docx_path)
    
    def _show_copy_toast(self):
        """显示复制成功提示（委托给剪贴板 helper）。"""
        show_copy_toast_for_app(self)
    
    def clear_all(self):
        """清空所有内容"""
        # 检查未保存的更改
        if not self._check_unsaved_changes():
            return
        
        self.input_text.delete("1.0", "end")
        self.current_file = None
        self._last_saved_content = ""
        self._content_modified = False
        self._update_title()
        self.on_text_change(None)
        self.update_status("✨ 已清空")

    def toggle_theme(self):
        self.theme_feature.toggle_theme()
    
    def update_status(self, message: str):
        """更新状态栏"""
        self.status_bar_feature.update_status(message)

    def _on_cursor_event(self, event=None):
        self._update_cursor_position()
        return None

    def _update_cursor_position(self):
        """更新状态栏的光标行/列"""
        try:
            tb = getattr(self.input_text, '_textbox', None)
            if tb is None:
                tb = getattr(self.input_text, 'text', None)
            if tb is None:
                return

            self.status_bar_feature.update_cursor_position(tb)
        except Exception:
            pass
    
    # ==================== 新增功能方法 ====================
    
    def toggle_preview(self):
        """切换预览显示/隐藏"""
        self.preview_visible = not self.preview_visible
        
        if self.preview_visible:
            # 显示预览
            self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
            try:
                if hasattr(self, 'hide_preview_btn') and self.hide_preview_btn is not None:
                    self.hide_preview_btn.configure(text="✕ 关闭预览")
            except Exception:
                pass
            # 调整列权重
            self.main_frame.grid_columnconfigure(0, weight=3)
            self.main_frame.grid_columnconfigure(1, weight=2)
            # 更新预览
            self.on_text_change(None)
            self.update_status("👁️ 预览已开启")
        else:
            # 隐藏预览
            self.preview_card.grid_forget()
            # 调整输入区域占满
            self.main_frame.grid_columnconfigure(0, weight=1)
            self.main_frame.grid_columnconfigure(1, weight=0)
            self.update_status("📝 纯编辑模式 - 按 Ctrl+P 或点击工具栏打开预览")

        self.header_styler.update_states()
    
    def toggle_sidebar(self):
        """切换侧边栏显示/隐藏"""
        self.sidebar_visible = not self.sidebar_visible
        
        if self.sidebar_visible:
            self.sidebar.pack(side="left", fill="y", padx=(0, 10), before=self.main_container.winfo_children()[1])
        else:
            self.sidebar.pack_forget()
        
        # 保存配置
        self.config['sidebar_visible'] = self.sidebar_visible
        save_config(self.config)

        self.header_styler.update_states()
    
    def change_font_size(self, delta: int):
        """调整字体大小"""
        current_size = self.config.get('font_size', 14)
        new_size = max(10, min(24, current_size + delta))
        
        if new_size != current_size:
            self.config['font_size'] = new_size
            save_config(self.config)
            
            # 更新编辑器字体
            if hasattr(self, 'input_editor'):
                self.input_editor.set_font_size(new_size)
            
            self.update_status(f"🔤 字体大小: {new_size}px")
    
    def show_search_dialog(self):
        """显示搜索替换对话框"""
        if self.search_dialog is None or not self.search_dialog.winfo_exists():
            # 获取实际的text widget
            text_widget = self.input_editor._textbox
            self.search_dialog = SearchReplaceDialog(self, text_widget)
        else:
            self.search_dialog.focus()

    def open_export_style_settings(self):
        try:
            dlg = getattr(self, 'export_style_dialog', None)
            if dlg is not None and dlg.winfo_exists():
                dlg.focus()
                return
        except Exception:
            pass

        try:
            self.export_style_dialog = ExportStyleSettingsDialog(self)
        except Exception:
            self.export_style_dialog = None
    
    def _setup_drag_drop(self):
        """设置拖拽导入支持"""
        try:
            # 尝试使用tkinterdnd2（如果安装了）
            from tkinterdnd2 import DND_FILES, TkinterDnD
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop)
        except ImportError:
            # 没有安装tkinterdnd2，使用简单的方式
            pass
    
    def _on_drop(self, event):
        """处理拖拽放置事件"""
        handle_drop_for_app(self, event)
    
    def _open_recent_file(self, file_path: str):
        """打开最近文件"""
        self.file_ops.open_recent_file(file_path)
    
    def _jump_to_line(self, line_number: int):
        """跳转到指定行"""
        try:
            # 设置光标位置
            index = f"{line_number}.0"
            self.input_text._textbox.see(index)
            self.input_text._textbox.mark_set("insert", index)
            self.input_text._textbox.focus()
        except Exception:
            pass
    
    # ==================== 文件保存功能 ====================
    
    def save_file(self):
        self.file_ops.save_file()

    def save_file_as(self):
        self.file_ops.save_file_as()
    
    def _check_unsaved_changes(self) -> bool:
        return self.file_ops.check_unsaved_changes()
    
    def _on_closing(self):
        """窗口关闭事件"""
        if self._check_unsaved_changes():
            # 保存窗口位置和大小
            self._save_window_geometry()
            self.destroy()
    
    def _update_title(self):
        """更新窗口标题"""
        base_title = "✨ Markdown → Word 转换器 by 一个好人"
        if self.current_file:
            filename = os.path.basename(self.current_file)
            modified = " *" if self._content_modified else ""
            self.title(f"{filename}{modified} - {base_title}")
        else:
            modified = " *" if self._content_modified else ""
            self.title(f"未命名{modified} - {base_title}")
    
    # ==================== 撤销重做 ====================
    
    def _undo(self):
        """撤销操作"""
        try:
            self.input_text._textbox.edit_undo()
            self.on_text_change(None)
        except tk.TclError:
            pass  # 没有可撤销的操作
    
    def _redo(self):
        """重做操作"""
        try:
            self.input_text._textbox.edit_redo()
            self.on_text_change(None)
        except tk.TclError:
            pass  # 没有可重做的操作
    
    # ==================== 窗口位置记忆 ====================
    
    def _save_window_geometry(self):
        """保存窗口位置和大小"""
        self.window_geometry_feature.save()
    
    def _restore_window_geometry(self):
        """恢复窗口位置和大小"""
        self.window_geometry_feature.restore()
    
    # ==================== 帮助菜单 ====================
    
    def show_help(self):
        self.help_dialog_feature.show()

    def _insert_table_template(self):
        self.insert_templates.insert_table_template()

    def _insert_link_template(self):
        self.insert_templates.insert_link_template()

    def _insert_image_template(self):
        self.insert_templates.insert_image_template()

    def _insert_math_template(self):
        self.insert_templates.insert_math_template()

    def _insert_code_template(self):
        self.insert_templates.insert_code_template()

    def _insert_hr_template(self):
        self.insert_templates.insert_hr_template()

    def _insert_task_template(self):
        self.insert_templates.insert_task_template()
    
    # ==================== 同步滚动 ====================
    
    def _on_editor_scroll(self, position: float):
        """编辑器滚动时同步预览区"""
        self.preview_sync.on_editor_scroll(position)
    
    # ==================== 自动保存 ====================

    def _check_auto_save_recovery(self):
        return self.auto_save_feature.check_recovery()

    def _clear_auto_save(self):
        self.auto_save_feature.clear()
    
    # ==================== 插入菜单 ====================
    
    def show_insert_menu(self, event=None):
        self.insert_templates.show_menu(event)
    
    

def main():
    """启动应用"""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
