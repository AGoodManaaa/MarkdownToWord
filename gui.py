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
    PDFExportFeature,
    PreviewZoomFeature,
    EditorZoomFeature,
    TabManagerFeature,
    StatisticsDetailFeature,
    UndoRedoFeature,
    # Phase 1 新增功能
    FocusModeFeature,
    ReadingModeFeature,
    TOCGeneratorFeature,
    WatermarkFeature,
    ThemeEditorFeature,
    TemplateManager,
    TemplateSelectorFeature,
    HeaderFooterFeature,
    # Phase 3 新增功能
    DocumentStatsFeature,
    GlobalSearchReplaceFeature,
    LinkCheckerFeature,
    SnippetLibraryFeature,
    BatchExportFeature,
    ChartEditorFeature,
    MindmapFeature,
    BibliographyFeature,
    VersionControlFeature,
    HTMLExportFeature,  # HTML 导出
    DiagramFeature,     # 图表支持
    # Phase 4 新增功能
    AIAssistantFeature,
    AutocompleteFeature,
    # Phase 5 新增功能 - OCR
    OCRFeature,
    # Phase 5 新增功能 - 数据库、协作
    DatabaseFeature,
    CollaborationFeature,
    # Phase 6 新增功能 - 用户体验优化
    KeyboardShortcutsFeature,
    FolderViewFeature,
)
# 新增编辑器增强模块
from ui.features.syntax_highlight import SyntaxHighlighter, HighlightTheme
from ui.features.smart_editing import SmartEditor
from ui.features.minimap import Minimap
from ui.features.code_folding import CodeFolding
from ui.features.split_screen import SplitScreenManager, FullScreenPreview, PrintPreview, SplitMode
from ui.icons import TOOLBAR_ICONS, icons, get_toolbar_icon, get_status_icon, get_message_icon
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
from ui.layout_builder import (
    build_header,
    build_main_content,
    build_sidebar_content,
    build_input_panel,
    build_preview_panel,
)


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
                try:
                    apply_window_icon(self)
                except Exception:
                    pass  # 图标加载失败不影响程序运行
        except Exception:
            pass  # 图标加载失败不影响程序运行
        
        # 加载配置
        self.config = load_config()
        self.product_mode = self.config.get('product_mode', 'converter')
        self.show_advanced_toolbar = bool(self.config.get('show_advanced_toolbar', False))

        # 模板管理器（共享配置，供导出与管理对话使用）
        self.template_manager = TemplateManager(self, self.config)

        self.file_ops = FileOpsFeature(self)
        self.theme_feature = ThemeFeature(self)

        self.preview_sync = PreviewSyncFeature(self)
        self.window_geometry_feature = WindowGeometryFeature(self)
        self.theme_feature.apply_mode('light')
        
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
        self.pdf_export_feature = PDFExportFeature(self)
        self.preview_zoom_feature = PreviewZoomFeature(self)
        self.editor_zoom_feature = EditorZoomFeature(self)
        self.tab_manager = TabManagerFeature(self)
        self.statistics_detail = StatisticsDetailFeature(self)
        
        # 将所有 UI 强相关的 Feature 提前初始化，确保在 _init_ui 之前全部实例化
        self.theme_editor = ThemeEditorFeature(self)
        self.reading_mode = ReadingModeFeature(self)
        self.toc_generator = TOCGeneratorFeature(self)
        self.watermark_feature = WatermarkFeature(self)
        self.template_selector = TemplateSelectorFeature(self)
        self.header_footer = HeaderFooterFeature(self)
        self.focus_mode = FocusModeFeature(self)  # 补上缺失的 focus_mode
        
        # Phase 3-6 功能也一并提前初始化，确保全局可用
        self._optional_feature_factories = {
            'document_stats_feature': lambda: DocumentStatsFeature(self),
            'global_search_replace': lambda: GlobalSearchReplaceFeature(self),
            'link_checker': lambda: LinkCheckerFeature(self),
            'snippet_library': lambda: SnippetLibraryFeature(self),
            'batch_export_feature': lambda: BatchExportFeature(self),
            'chart_editor': lambda: ChartEditorFeature(self),
            'mindmap_feature': lambda: MindmapFeature(self),
            'bibliography_feature': lambda: BibliographyFeature(self),
            'version_control': lambda: VersionControlFeature(self),
            'html_export_feature': lambda: HTMLExportFeature(self),
            'diagram_feature': lambda: DiagramFeature(self),
            'ai_assistant': lambda: AIAssistantFeature(self),
            'ocr_feature': lambda: OCRFeature(self),
            'database_feature': lambda: DatabaseFeature(self),
            'collaboration_feature': lambda: CollaborationFeature(self),
            'keyboard_shortcuts': lambda: KeyboardShortcutsFeature(self),
            'folder_view_feature': lambda: FolderViewFeature(self),
        }
        for attr_name in self._optional_feature_factories:
            setattr(self, attr_name, None)

        # 彻底移除自定义撤销系统的引用，完全回归原生
        # if hasattr(self, 'undo_redo'): del self.undo_redo
        
        # 初始化 UI
        self._init_ui()
        
        # 初始化需要在UI创建后加载的功能
        self.autocomplete_feature = AutocompleteFeature(self)
        
        # 初始化编辑器增强功能（在UI创建后）
        self._init_editor_enhancements()
        
        # 内容修改标记
        self._content_modified = False
        self._last_saved_content = ""
        self._last_content_snapshot = None

    def _ensure_optional_feature(self, attr_name: str, display_name: str = ""):
        feature = getattr(self, attr_name, None)
        if feature is not None:
            return feature

        factory = self._optional_feature_factories.get(attr_name)
        if factory is None:
            return None

        try:
            feature = factory()
            setattr(self, attr_name, feature)
            return feature
        except Exception as exc:
            print(f"Failed to initialize optional feature '{attr_name}': {exc}")
            if hasattr(self, 'status_bar_feature'):
                try:
                    self.update_status(f"{display_name or attr_name} 鍔犺浇澶辫触")
                except Exception:
                    pass
            return None

    def _init_ui(self):
        """初始化用户界面"""
        # 构建界面
        self._create_header()
        self._create_status_bar()  # 先创建状态栏
        self._create_main_content()  # 再创建主内容（包含_insert_example调用）
        
        # 建立内容变化的监听 (彻底解决自动刷新问题)
        # 通过监听 Text 内部的 <<Modified>> 事件，使用 add=True 避免干扰其他功能（如撤销系统）
        if hasattr(self, 'input_text') and hasattr(self.input_text, '_textbox'):
            self.input_text._textbox.bind('<<Modified>>', self._on_text_modified_event, add=True)
        self.header_styler.update_states()
        
        # 绑定快捷键
        self.bind('<Control-o>', lambda e: self.open_file())
        self.bind('<Control-s>', lambda e: self.save_file())  # 保存源文件
        self.bind('<Control-Shift-s>', lambda e: self.export_to_word())  # 导出（打开导出选项）
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
        self.bind('<F11>', lambda e: self.focus_mode.toggle())  # 专注模式
        self.bind('<F12>', lambda e: self.reading_mode.toggle())  # 阅读模式
        self.bind('<Control-F11>', lambda e: self.toggle_fullscreen_preview())  # 全屏预览
        self.bind('<Control-Shift-p>', lambda e: self.show_print_preview())  # 打印预览
        if self.show_advanced_toolbar:
            self.bind('<Control-Shift-o>', lambda e: self.show_ocr())  # OCR 功能
            self.bind('<Control-Shift-d>', lambda e: self.show_database())  # 文档库
            self.bind('<Control-Alt-c>', lambda e: self.show_collaboration())  # 协作
        
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
        
        # 恢复预览缩放比例
        self.preview_zoom_feature.restore_scale()
        
        # 恢复编辑区缩放比例
        self.editor_zoom_feature.restore_scale()
        
        # 设置撤销/重做系统（必须在编辑器创建后）
        self.after(100, self._setup_undo_system)  # 延迟100ms确保编辑器完全初始化
    
        self.config['split_mode'] = SplitMode.HORIZONTAL.value
        self.config['split_ratio'] = 0.5
        self.preview_visible = True
        self.after(120, self._enforce_initial_split_layout)

    def _setup_undo_system(self):
        """设置极致精细化撤销系统"""
        try:
            if hasattr(self, 'input_text') and hasattr(self.input_text, '_textbox'):
                t = self.input_text._textbox
                
                # 1. 彻底清除任何自定义干扰
                for attr in ('undo_manager', '_original_insert', '_original_delete'):
                    if hasattr(t, attr): delattr(t, attr)
                
                # 2. 启用原生撤销，禁用自动分隔，由编辑器精细控制
                t.configure(undo=True, autoseparators=False, maxundo=5000)
                
                # 3. 建立初始锚点
                t.edit_reset()
                t.edit_separator()
                print("✅ 极致精细化撤销系统已就绪")
        except Exception as e:
            print(f"Undo system error: {e}")
    
    def _init_editor_enhancements(self):
        """初始化编辑器增强功能"""
        try:
            # 获取底层文本组件
            text_widget = self.input_text._textbox if hasattr(self.input_text, '_textbox') else self.input_text
            
            # 1. 语法高亮
            self.syntax_highlighter = SyntaxHighlighter(text_widget)
            print("✅ 语法高亮已启用")
            
            # 2. 智能编辑（缩进、括号匹配）
            self.smart_editor = SmartEditor(text_widget)
            self.input_editor.smart_editor_ref = self.smart_editor # 建立双向引用用于更新
            print("✅ 智能编辑已启用")
            
            # 3. 迷你地图（嵌入到编辑区内部右侧，VS Code 风格）
            # 使用 text_frame 作为父容器，这样迷你地图会在文本区域内部
            minimap_parent = self.input_editor.text_frame if hasattr(self.input_editor, 'text_frame') else self.input_card
            self.minimap = Minimap(text_widget, minimap_parent)
            self.minimap_visible = self.config.get('minimap_visible', True)
            if self.minimap_visible:
                self.minimap.show()
            print("✅ 迷你地图已启用 (VS Code 风格)")
            
            # 4. 代码折叠
            line_canvas = None
            if hasattr(self.input_editor, 'line_numbers') and hasattr(self.input_editor.line_numbers, 'canvas'):
                line_canvas = self.input_editor.line_numbers.canvas
            self.code_folding = CodeFolding(text_widget, line_canvas)
            print("✅ 代码折叠已启用")
            
            # 5. 预览主题管理器
            # 6. 分屏管理器
            self.split_screen = SplitScreenManager(self)
            print("✅ 分屏管理器已启用")
            self.config['split_mode'] = SplitMode.HORIZONTAL.value
            self.config['split_ratio'] = 0.5
            self.preview_visible = True
            self.after(50, self._enforce_initial_split_layout)
            
            # 7. 全屏预览
            self.fullscreen_preview = FullScreenPreview(self)
            print("✅ 全屏预览已启用")
            
            # 8. 打印预览
            self.print_preview = PrintPreview(self)
            print("✅ 打印预览已启用")
            
        except Exception as e:
            print(f"⚠️ 编辑器增强功能初始化失败: {e}")
            import traceback
            traceback.print_exc()
    
    def _legacy_create_header(self):
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
        
        # 工具按钮：默认聚焦转换主链路，高级能力按需显示
        tools = [
            (get_toolbar_icon("open"), "打开", self.open_file, "Ctrl+O"),
            (get_toolbar_icon("save"), "保存", self.save_file, "Ctrl+S"),
            (get_toolbar_icon("format"), "规范化", self.format_markdown, "Ctrl+Shift+F"),
            (get_toolbar_icon("search"), "搜索", self.show_search_dialog, "Ctrl+F"),
            (get_toolbar_icon("preview"), "预览", self.toggle_preview, "Ctrl+P"),
            ("↔️", "左右分屏", self.set_split_horizontal, ""),
            ("↕️", "上下分屏", self.set_split_vertical, ""),
            ("📝", "仅编辑", self.set_editor_only, ""),
            ("👁", "仅预览", self.set_preview_only, ""),
            (get_toolbar_icon("export"), "导出", self.export_to_word, "Ctrl+Shift+S"),
            (get_toolbar_icon("pdf"), "PDF", self.export_to_pdf, ""),
            ("🌐", "HTML", self.show_html_export, ""),
            (get_toolbar_icon("batch"), "批量导出", self.show_batch_export, ""),
        ]
        if self.show_advanced_toolbar:
            tools.extend([
                (get_toolbar_icon("ocr"), "OCR", self.show_ocr, "Ctrl+Shift+O"),
                (get_toolbar_icon("ai"), "AI助手", self.show_ai_assistant, "Ctrl+I"),
                (get_toolbar_icon("chart"), "图表", self.show_chart_editor, ""),
                (get_toolbar_icon("mindmap"), "导图", self.show_mindmap, ""),
                (get_toolbar_icon("bibliography"), "文献", self.show_bibliography, ""),
                (get_toolbar_icon("version"), "版本", self.show_version_control, ""),
                (get_toolbar_icon("link"), "链接", self.show_link_checker, ""),
                (get_toolbar_icon("database"), "文档库", self.show_database, "Ctrl+Shift+D"),
                (get_toolbar_icon("collab"), "协作", self.show_collaboration, "Ctrl+Alt+C"),
            ])
        
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
                font=("Segoe UI Emoji", 16),
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
            text=get_toolbar_icon("insert"),
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
            text=get_toolbar_icon("sidebar"),
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
            text=get_toolbar_icon("font_minus"),
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
            text=get_toolbar_icon("font_plus"),
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
            text=(get_toolbar_icon("theme_light") if ctk.get_appearance_mode() == "Dark" else get_toolbar_icon("theme_dark")),
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
        self.tooltip.add_tooltip(self.theme_btn, "切换亮/暗主题")
        
        # 自定义主题
        self.theme_editor_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("theme_editor"),
            command=self.theme_editor.show_editor,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.theme_editor_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.theme_editor_btn)
        self.tooltip.add_tooltip(self.theme_editor_btn, "自定义主题编辑器")
        
        # Phase 1: 专注模式按钮
        self.focus_mode_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("focus"),
            command=self.focus_mode.toggle,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.focus_mode_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.focus_mode_btn)
        self.tooltip.add_tooltip(self.focus_mode_btn, "专注模式\nF11")
        
        # Phase 1: 阅读模式按钮
        self.reading_mode_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("reading"),
            command=self.reading_mode.toggle,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.reading_mode_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.reading_mode_btn)
        self.tooltip.add_tooltip(self.reading_mode_btn, "阅读模式\nF12")
        
        # 迷你地图切换按钮
        self.minimap_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("minimap"),
            command=self.toggle_minimap,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.minimap_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.minimap_btn)
        self.tooltip.add_tooltip(self.minimap_btn, "迷你地图\n文档缩略导航")
        
        # 分屏模式按钮
        self.split_mode_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("split"),
            command=self.toggle_split_mode,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.split_mode_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.split_mode_btn)
        self.tooltip.add_tooltip(self.split_mode_btn, "切换分屏模式\n左右/上下/仅编辑/仅预览")
        
        # 全屏预览按钮
        self.fullscreen_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("fullscreen"),
            command=self.toggle_fullscreen_preview,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34
        )
        self.fullscreen_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.fullscreen_btn)
        self.tooltip.add_tooltip(self.fullscreen_btn, "全屏预览\nCtrl+F11")

        self.export_style_header_btn = ctk.CTkButton(
            btn_frame,
            text=get_toolbar_icon("settings"),
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
        
        # 快捷键设置按钮
        self.shortcuts_btn = ctk.CTkButton(
            btn_frame,
            text="⌨️",
            command=self.show_keyboard_shortcuts,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS['primary_hover'],
            corner_radius=10,
            width=38,
            height=34,
        )
        self.shortcuts_btn.pack(side="left", padx=3)
        self._header_default_buttons.append(self.shortcuts_btn)
        self.tooltip.add_tooltip(self.shortcuts_btn, "快捷键设置")

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
        
        # 右侧主编辑区（包含标签栏和编辑/预览区）
        self.right_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.right_container.pack(side="left", fill="both", expand=True)
        
        # 标签栏
        self.tab_bar = self.tab_manager.create_tab_bar(self.right_container)
        self.tab_bar.pack(fill="x", pady=(0, 4))
        
        # 编辑/预览区
        self.main_frame = ctk.CTkFrame(self.right_container, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True)
        
        # 使用 tk.PanedWindow 实现可拖动分栏
        self.paned_window = tk.PanedWindow(
            self.main_frame, 
            orient=tk.HORIZONTAL, 
            bg=COLORS['bg_light'], 
            bd=0, 
            sashwidth=4,
            sashrelief='flat'
        )
        self.paned_window.pack(fill="both", expand=True)
        self.paned_window.bind("<ButtonRelease-1>", self._on_split_drag_end)
        
        # ===== 左侧：输入区域 =====
        self._create_input_panel(self.paned_window)
        self.paned_window.add(self.input_card, stretch="always")
        
        # ===== 右侧：预览区域 =====
        self._create_preview_panel(self.paned_window)
        self.paned_window.add(self.preview_card, stretch="always")
        
        # 初始权重设置：左右 50/50（基于 paned_window 实际宽度）
        def _set_half_half():
            try:
                w = self.paned_window.winfo_width()
                if w <= 0:
                    self.after(100, _set_half_half)
                    return
                ratio = float(self.config.get('split_ratio', 0.5))
                self.paned_window.sash_place(0, int(w * ratio), 0)
            except Exception:
                pass
        self.after(150, _set_half_half)
        
        # 插入示例文本（在所有组件创建完成后）
        self._insert_example()
    
    def _create_sidebar_content(self):
        """创建侧边栏内容"""
        # 文件夹视图
        self.folder_view = None
        folder_feature = self._ensure_optional_feature('folder_view_feature', '文件夹视图')
        if folder_feature:
            try:
                self.folder_view = folder_feature.create_view(self.sidebar)
                self.folder_view.pack(fill="both", expand=True, pady=(0, 5))
            except Exception:
                self.folder_view = None
        if self.folder_view is None:
            folder_placeholder = ctk.CTkFrame(self.sidebar, fg_color="transparent")
            folder_placeholder.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(
                folder_placeholder,
                text="文件夹视图未启用",
                text_color=COLORS['text_secondary'],
                anchor="w",
            ).pack(fill="x", pady=(6, 2))
            ctk.CTkButton(
                folder_placeholder,
                text="打开文件夹",
                height=30,
                command=self.open_folder,
            ).pack(fill="x")
        
        # 分隔线
        separator1 = ctk.CTkFrame(self.sidebar, height=1, fg_color=COLORS['border'])
        separator1.pack(fill="x", padx=15, pady=5)
        
        # 大纲视图
        self.outline_view = OutlineView(
            self.sidebar,
            on_heading_click=self._jump_to_line
        )
        self.outline_view.pack(fill="both", expand=True, pady=(0, 10))
        
        # 分隔线
        separator2 = ctk.CTkFrame(self.sidebar, height=1, fg_color=COLORS['border'])
        separator2.pack(fill="x", padx=15, pady=5)
        
        # 最近文件
        self.recent_files_view = RecentFilesView(
            self.sidebar,
            on_file_click=self._open_recent_file
        )
        self.recent_files_view.pack(fill="both", expand=True)
    
    def _create_input_panel(self, parent):
        """创建输入面板 - 带行号"""
        self.input_card = ModernCard(parent)
        # self.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8)) # PanedWindow.add handles this
        
        # 工具栏 - 紧凑布局，紧贴文本框
        toolbar = ctk.CTkFrame(self.input_card, fg_color="transparent", height=26)
        toolbar.pack(fill="x", padx=6, pady=(6, 0))
        toolbar.pack_propagate(False)  # 保持固定高度
        
        # 编辑区缩放控件（放在工具栏右侧）
        editor_zoom_controls = self.editor_zoom_feature.create_controls(toolbar)
        editor_zoom_controls.pack(side="right", padx=(4, 0))
        
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
        # 注意：必须绑定到内部的 text 组件，因为 Frame 不会接收键盘事件，导致自动刷新失效
        if hasattr(self.input_editor, 'text'):
            self.input_editor.text.bind('<KeyRelease>', self._on_text_change_debounced, add="+")
        else:
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
        # self.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0)) # PanedWindow.add handles this
        
        # 顶部控件栏
        top_frame = ctk.CTkFrame(self.preview_card, fg_color="transparent", height=30)
        top_frame.pack(fill="x", padx=10, pady=(4, 0))
        
        # 右侧：缩放控件
        self.preview_zoom_controls = self.preview_zoom_feature.create_controls(top_frame)
        self.preview_zoom_controls.pack(side="right", padx=(0, 4))
        
        # 预览组件
        self.preview = MarkdownPreview(
            self.preview_card, 
            on_content_change=self._on_preview_change, 
            app=self,
            on_scroll=self._on_preview_scroll
        )
        self.preview.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        
        # 预览区域提示（只读、双击跳转、Ctrl+C）
        hint_frame = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        hint_frame.pack(fill="x", padx=10, pady=(0, 6))
        hint_label = ctk.CTkLabel(
            hint_frame,
            text="👁️ 预览区为只读，支持 Ctrl+C 复制，双击任意段落可跳转到源码；滚动同步可在编辑器/预览滚动时自动保持一致。",
            anchor="w",
            text_color=COLORS.get('text_secondary', '#6b7280'),
            font=ctk.CTkFont(size=11)
        )
        hint_label.pack(side="left", fill="x")
        
        # 设置预览区双击跳转回调
        self.preview.set_jump_callback(self._jump_to_line)
        
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
        
        # 手动刷新按钮
        self.refresh_preview_btn = ModernButton(
            right_group,
            text="🔄 刷新",
            command=self.refresh_preview,
            style="ghost",
            width=80
        )
        self.refresh_preview_btn.pack(side="right", padx=(0, 6))
    
    def _on_split_drag_end(self, event=None):
        """拖拽分隔条后保存比例，保持下次启动一致。"""
        try:
            if not hasattr(self, 'paned_window') or len(self.paned_window.panes()) < 2:
                return
            orient = self.paned_window.cget("orient")
            total_w = max(1, self.paned_window.winfo_width())
            total_h = max(1, self.paned_window.winfo_height())
            try:
                sash_x, sash_y = self.paned_window.sash_coord(0)
            except Exception:
                sash_x, sash_y = 0, 0
            if str(orient) == str(tk.VERTICAL):
                ratio = sash_y / float(total_h)
            else:
                ratio = sash_x / float(total_w)
            ratio = min(0.85, max(0.15, ratio))
            self.config['split_ratio'] = ratio
            save_config(self.config)
        except Exception:
            pass
    
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_bar_feature.create()
        self.status_bar = self.status_bar_feature.frame
        self.status_label = self.status_bar_feature.status_label
        self.word_count_label = self.status_bar_feature.word_count_label
        self.cursor_pos_label = self.status_bar_feature.cursor_pos_label
        
        # 绑定状态栏点击事件显示详细统计
        self.statistics_detail.bind_status_bar_click()
    
    def _insert_example(self):
        """插入示例Markdown"""
        insert_example_if_empty_for_app(self)
        # 强制在示例文本后插入一个基准点，确保后续修改是分步的
        try:
            if hasattr(self.input_text, '_textbox'):
                self.input_text._textbox.edit_separator()
                print("✅ 已建立撤销基准点")
        except Exception:
            pass
    
    def insert_text(self, text: str):
        """在光标位置插入文本"""
        self.input_text.insert("insert", text)
        self.on_text_change(None)
    
    def _get_editor_content(self) -> str:
        try:
            return self.input_text.get("1.0", "end-1c")
        except Exception:
            return ""

    def _update_welcome_state(self) -> None:
        panel = getattr(self, "welcome_panel", None)
        editor = getattr(self, "input_editor", None)
        if panel is None or editor is None:
            return

        has_content = bool((self._get_editor_content() or "").strip())
        visible = bool(getattr(panel, "_visible", True))

        if has_content and visible:
            try:
                panel.pack_forget()
                panel._visible = False
            except Exception:
                pass
        elif (not has_content) and (not visible):
            try:
                panel.pack(fill="x", padx=6, pady=(8, 4), before=editor)
                panel._visible = True
            except Exception:
                pass

    def paste_from_clipboard(self) -> None:
        try:
            clipboard_text = self.clipboard_get()
        except Exception:
            clipboard_text = ""

        if not clipboard_text:
            self.update_status("剪贴板里没有可粘贴的文本")
            return

        try:
            if not self._get_editor_content().strip():
                self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", clipboard_text)
            self.on_text_change(None)
            self.update_status("已从剪贴板导入内容")
        except Exception as exc:
            self.update_status(f"粘贴失败: {exc}")

    def load_welcome_sample(self) -> None:
        sample = """# 项目状态周报

## 本周进展

- 完成 Markdown 到 Word 的导出优化
- 新增模板选择与导出前检查
- 修复图片路径和表格格式问题

## 风险项

| 项目 | 风险 | 处理建议 |
| --- | --- | --- |
| 图片 | 路径失效 | 导出前检查 |
| 表格 | 列数不一致 | 统一表头格式 |

## 公式

$$
E = mc^2
$$
"""
        try:
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", sample)
            self.on_text_change(None)
            self.update_status("已插入示例内容")
        except Exception as exc:
            self.update_status(f"插入示例失败: {exc}")

    def _on_text_change_debounced(self, event):
        self.preview_sync.on_text_change_debounced(event)
    
    def on_text_change(self, event):
        self._update_welcome_state()
        self.preview_sync.on_text_change(event)
    
    def _on_preview_change(self, markdown_text: str):
        self.preview_sync.on_preview_change(markdown_text)
    
    def _on_preview_scroll(self, position: float):
        """预览区滚动时同步编辑器"""
        self.preview_sync.on_preview_scroll(position)

    def open_file(self):
        self.file_ops.open_file()

    def export_to_word(self):
        """导出为Word文档（委托给导出 helper）。"""
        export_to_word_for_app(self)

    def export_to_pdf(self):
        """导出为PDF文档。"""
        self.pdf_export_feature.export_to_pdf()

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
    
    def show_document_stats(self):
        """显示文档统计分析。"""
        try:
            feature = self._ensure_optional_feature('document_stats_feature', '文档统计')
            if feature:
                feature.show_stats()
        except Exception:
            pass
    
    def show_link_checker(self):
        """显示链接检查器。"""
        try:
            feature = self._ensure_optional_feature('link_checker', '链接检查')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_snippet_library(self):
        """显示片段库。"""
        try:
            feature = self._ensure_optional_feature('snippet_library', '片段库')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_batch_export(self):
        """显示批量导出。"""
        try:
            feature = self._ensure_optional_feature('batch_export_feature', '批量导出')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_html_export(self):
        """显示 HTML 导出对话框。"""
        try:
            feature = self._ensure_optional_feature('html_export_feature', 'HTML 导出')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_chart_editor(self):
        """显示图表编辑器。"""
        try:
            feature = self._ensure_optional_feature('chart_editor', '图表')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_mindmap(self):
        """显示思维导图转换器。"""
        try:
            feature = self._ensure_optional_feature('mindmap_feature', '导图')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_bibliography(self):
        """显示文献引用管理。"""
        try:
            feature = self._ensure_optional_feature('bibliography_feature', '文献')
            if feature:
                feature.show_dialog()
        except Exception:
            pass
    
    def show_version_control(self):
        """显示版本历史。"""
        try:
            feature = self._ensure_optional_feature('version_control', '版本历史')
            if feature:
                feature.show_dialog()
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
        """更新状态栏的光标行/列及选中统计"""
        try:
            tb = getattr(self.input_text, '_textbox', None)
            if tb is None:
                tb = getattr(self.input_text, 'text', None)
            if tb is None:
                return

            # 更新行列信息
            self.status_bar_feature.update_cursor_position(tb)
            
            # 更新选中统计
            try:
                selected_text = tb.get(tk.SEL_FIRST, tk.SEL_LAST)
                selected_count = len(selected_text.replace('\n', '').replace(' ', '').replace('\t', ''))
            except tk.TclError:
                selected_count = 0
                
            content = tb.get("1.0", "end-1c")
            if hasattr(self, 'statistics_detail'):
                self.statistics_detail.update_status_bar(content, selected_count)
            else:
                self.status_bar_feature.update_counts(content, selected_count)
        except Exception:
            pass
    
    # ==================== 新增功能方法 ====================
    
    def toggle_preview(self):
        """切换预览显示/隐藏"""
        self.preview_visible = not self.preview_visible
        
        if self.preview_visible:
            # 显示预览
            self.paned_window.add(self.preview_card)
            try:
                if hasattr(self, 'hide_preview_btn') and self.hide_preview_btn is not None:
                    self.hide_preview_btn.configure(text="✕ 关闭预览")
            except Exception:
                pass
            # 更新预览
            self.on_text_change(None)
            self.update_status("👁️ 预览已开启")
        else:
            # 隐藏预览
            self.paned_window.forget(self.preview_card)
            self.update_status("� 纯编辑模式 - 按 Ctrl+P 或点击工具栏打开预览")

        self.header_styler.update_states()
    
    def refresh_preview(self):
        """手动刷新预览内容"""
        try:
            if hasattr(self, 'input_text') and hasattr(self, 'preview'):
                content = self.input_text.get("1.0", "end-1c")
                self._set_preview_status("渲染中...", None)
                self.preview.set_updating(True)
                self.preview.update_preview(content)
                self.preview.set_updating(False)
                self._set_preview_status("预览就绪", None)
                self.update_status("🔄 预览已刷新")
        except Exception:
            self._set_preview_status("预览失败", None)
    
    def _toggle_scroll_sync(self):
        """切换滚动同步开关"""
        try:
            enabled = self.preview_sync.toggle_scroll_sync()
        except Exception:
            enabled = getattr(self.preview_sync, '_scroll_sync_enabled', True)
        self._update_scroll_sync_btn()
        try:
            self.update_status("🔗 滚动同步已开启" if enabled else "📴 滚动同步已关闭")
        except Exception:
            pass
    
    def _update_scroll_sync_btn(self):
        """根据状态更新同步按钮文案"""
        try:
            enabled = getattr(self.preview_sync, '_scroll_sync_enabled', True)
            self.scroll_sync_btn.configure(text="同步开" if enabled else "同步关")
        except Exception:
            pass
    
    def _set_preview_status(self, text: str, color=None):
        """更新预览状态标签"""
        try:
            if hasattr(self, 'preview_status_label'):
                kwargs = {"text": text}
                if color is not None:
                    kwargs["text_color"] = color
                self.preview_status_label.configure(**kwargs)
        except Exception:
            pass
    
    def toggle_sidebar(self):
        """切换侧边栏显示/隐藏"""
        self.sidebar_visible = not self.sidebar_visible
        
        if self.sidebar_visible:
            self.sidebar.pack(side="left", fill="y", padx=(0, 10), before=self.right_container)
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
            if hasattr(self, 'syntax_highlighter'):
                try:
                    self.syntax_highlighter.refresh_styles()
                except Exception:
                    pass
            
            self.update_status(f"🔤 字体大小: {new_size}px")
    
    def show_ai_assistant(self):
        """显示 AI 写作助手"""
        feature = self._ensure_optional_feature('ai_assistant', 'AI 助手')
        if feature:
            feature.show_dialog()
    
    def show_ocr(self):
        """显示 OCR 图片转 Markdown 对话框"""
        feature = self._ensure_optional_feature('ocr_feature', 'OCR')
        if feature:
            feature.show_dialog()
    
    def show_database(self):
        """显示 Markdown 数据库/文档库功能"""
        feature = self._ensure_optional_feature('database_feature', '文档库')
        if feature:
            feature.show_vault_selector()
    
    def show_collaboration(self):
        """显示实时协作功能"""
        feature = self._ensure_optional_feature('collaboration_feature', '协作')
        if feature:
            feature.show_dialog()
    
    def show_keyboard_shortcuts(self):
        """显示快捷键设置"""
        feature = self._ensure_optional_feature('keyboard_shortcuts', '快捷键')
        if feature:
            feature.show_dialog()
    
    def open_folder(self):
        """打开文件夹"""
        feature = self._ensure_optional_feature('folder_view_feature', '文件夹视图')
        if feature:
            feature.open_folder()
    
    def toggle_minimap(self):
        """切换迷你地图显示"""
        if hasattr(self, 'minimap'):
            self.minimap_visible = not self.minimap_visible
            if self.minimap_visible:
                self.minimap.show()
                self.update_status("🗺️ 迷你地图已开启")
            else:
                self.minimap.hide()
                self.update_status("🗺️ 迷你地图已关闭")
            
            # 保存配置
            self.config['minimap_visible'] = self.minimap_visible
            save_config(self.config)
    
    def toggle_syntax_highlight(self, enabled: bool = None):
        """切换语法高亮"""
        if hasattr(self, 'syntax_highlighter'):
            if enabled is None:
                enabled = not self.syntax_highlighter._enabled
            
            if enabled:
                self.syntax_highlighter.enable()
                self.update_status("✨ 语法高亮已开启")
            else:
                self.syntax_highlighter.disable()
                self.update_status("✨ 语法高亮已关闭")
    
    def toggle_smart_editing(self, enabled: bool = None):
        """切换智能编辑"""
        if hasattr(self, 'smart_editor'):
            if enabled is None:
                # 切换状态
                if self.smart_editor.smart_indent._enabled:
                    self.smart_editor.disable_all()
                    self.update_status("📝 智能编辑已关闭")
                else:
                    self.smart_editor.enable_all()
                    self.update_status("📝 智能编辑已开启")
            elif enabled:
                self.smart_editor.enable_all()
            else:
                self.smart_editor.disable_all()
    
    def toggle_code_folding(self, enabled: bool = None):
        """切换代码折叠"""
        if hasattr(self, 'code_folding'):
            if enabled is None:
                enabled = not self.code_folding._enabled
            
            if enabled:
                self.code_folding.enable()
                self.update_status("📁 代码折叠已开启")
            else:
                self.code_folding.disable()
                self.update_status("📁 代码折叠已关闭")
    
    def fold_all_code(self):
        """折叠所有代码块"""
        if hasattr(self, 'code_folding'):
            self.code_folding.fold_all()
            self.update_status("📁 已折叠所有代码块")
    
    def unfold_all_code(self):
        """展开所有代码块"""
        if hasattr(self, 'code_folding'):
            self.code_folding.unfold_all()
            self.update_status("📂 已展开所有代码块")
    
    def set_split_horizontal(self):
        """设置左右分屏"""
        if hasattr(self, 'split_screen'):
            self.split_screen.set_horizontal()

    def _enforce_initial_split_layout(self):
        """启动后多次校准左右分屏，避免初始化时序导致比例失真"""
        self.config['split_mode'] = SplitMode.HORIZONTAL.value
        self.config['split_ratio'] = 0.5
        self.preview_visible = True

        if hasattr(self, 'split_screen'):
            self.split_screen.set_horizontal()

        def _place_half():
            try:
                if not hasattr(self, 'paned_window'):
                    return
                self.paned_window.update_idletasks()
                width = max(1, self.paned_window.winfo_width())
                self.paned_window.sash_place(0, int(width * 0.5), 0)
            except Exception:
                pass

        for delay in (0, 120, 260, 520):
            self.after(delay, _place_half)
    
    def set_split_vertical(self):
        """设置上下分屏"""
        if hasattr(self, 'split_screen'):
            self.split_screen.set_vertical()
    
    def set_editor_only(self):
        """设置仅编辑器模式"""
        if hasattr(self, 'split_screen'):
            self.split_screen.set_editor_only()
    
    def set_preview_only(self):
        """设置仅预览模式"""
        if hasattr(self, 'split_screen'):
            self.split_screen.set_preview_only()
    
    def toggle_split_mode(self):
        """循环切换分屏模式"""
        if hasattr(self, 'split_screen'):
            self.split_screen.toggle_mode()
    
    def toggle_fullscreen_preview(self):
        """切换全屏预览"""
        if hasattr(self, 'fullscreen_preview'):
            self.fullscreen_preview.toggle()
    
    def show_print_preview(self):
        """显示打印预览"""
        if hasattr(self, 'print_preview'):
            self.print_preview.show()
    
    def show_search_dialog(self):
        """显示搜索替换对话框 - 优先使用内联搜索悬浮层"""
        if hasattr(self, 'input_editor'):
            # 获取选中文本作为初始搜索词
            try:
                textbox = self.input_editor._textbox
                initial_text = textbox.get(tk.SEL_FIRST, tk.SEL_LAST)
            except:
                initial_text = ""
            self.input_editor.show_search_overlay(initial_text)
            self.update_status("🔍 已打开内联搜索 - 输入关键词实时高亮")
            return

        # 降级使用旧的全局搜索替换功能
        global_search = self._ensure_optional_feature('global_search_replace', '全局搜索替换')
        if global_search:
            global_search.show_dialog()
        elif self.search_dialog is None or not self.search_dialog.winfo_exists():
            # 回退到旧版对话框
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
        """跳转到指定行并高亮显示"""
        try:
            # 设置光标位置
            index = f"{line_number}.0"
            self.input_text._textbox.see(index)
            self.input_text._textbox.mark_set("insert", index)
            self.input_text._textbox.focus()
            
            # 高亮显示跳转的行（短暂闪烁效果）
            self._highlight_editor_line(line_number)
        except Exception:
            pass
    
    def _highlight_editor_line(self, line_number: int):
        """短暂高亮编辑器中的行"""
        try:
            text_widget = self.input_text._textbox
            
            # 配置高亮标签
            text_widget.tag_configure('jump_highlight', background='#fef3c7')
            
            # 添加高亮
            text_widget.tag_add('jump_highlight', f"{line_number}.0", f"{line_number}.end")
            
            # 300ms 后移除高亮
            self.after(300, lambda: self._remove_editor_highlight(line_number))
        except Exception:
            pass
    
    def _remove_editor_highlight(self, line_number: int):
        """移除编辑器行高亮"""
        try:
            self.input_text._textbox.tag_remove('jump_highlight', f"{line_number}.0", f"{line_number}.end")
        except Exception:
            pass

    # ==================== 文件保存功能 ====================
    
    def new_file(self):
        """新建文件"""
        # 使用标签管理器新建文件
        self.tab_manager.new_tab()
    
    def open_file(self):
        """打开文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                
                # 在新标签页中打开
                self.tab_manager.open_file_in_tab(file_path, content)
                self.file_ops.add_recent_file(file_path)
                
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def save_file(self):
        """保存文件"""
        # 使用标签管理器保存当前标签
        current_tab = self.tab_manager.get_active_tab()
        if current_tab:
            # 更新当前内容到标签数据
            current_tab.content = self.input_text.get("1.0", "end-1c")
            if self.tab_manager._save_tab(current_tab):
                self.update_status(f"已保存: {os.path.basename(current_tab.file_path)}")
    
    def save_file_as(self):
        """文件另存为"""
        current_tab = self.tab_manager.get_active_tab()
        if current_tab:
            current_tab.content = self.input_text.get("1.0", "end-1c")
            if self.tab_manager.save_tab_as(current_tab):
                self.update_status(f"已另存为: {os.path.basename(current_tab.file_path)}")

    def _check_unsaved_changes(self) -> bool:
        """检查所有标签页是否有未保存的更改"""
        return self.tab_manager.check_all_tabs_unsaved_changes()
    
    def _on_text_modified_event(self, event=None):
        """中心化的内容变动分发器：处理预览、大纲、高亮、行号和修改标记"""
        try:
            # 检查是否是真的修改
            if self.input_text._textbox.edit_modified():
                # 1. 触发防抖后的预览和大纲更新
                self._on_text_change_debounced(None)
                
                # 2. 手动驱动子模块更新
                try:
                    if hasattr(self, 'syntax_highlighter'):
                        self.syntax_highlighter._on_key_release()
                    if hasattr(self.input_text, 'line_numbers'):
                        # 确保 line_numbers 是 LineNumbers 实例
                        ln = getattr(self.input_text, 'line_numbers', None)
                        if ln and hasattr(ln, '_update'):
                            ln._update()
                    if hasattr(self, 'minimap'):
                        self.minimap._schedule_update()
                except Exception as e:
                    print(f"Submodule update error: {e}")

                # 3. 标记内容已修改（用于标题显示 * 号）
                if not self._content_modified:
                    self._content_modified = True
                    self._update_title()
                
                # 4. 重要：重置标记以允许下一次触发
                self.input_text._textbox.edit_modified(False)
        except Exception as e:
            print(f"Centralized refresh error: {e}")
    
    def _on_closing(self):
        """窗口关闭事件"""
        try:
            if self._check_unsaved_changes():
                # 保存窗口位置和大小
                try:
                    self._save_window_geometry()
                except:
                    pass
                
                # 强制退出程序
                try:
                    self.quit()     # 停止 mainloop
                    self.destroy()  # 销毁窗口
                except:
                    pass
                
                import sys
                sys.exit(0)  # 强制结束进程
                
        except Exception as e:
            print(f"Closing error: {e}")
            # 最后的重试
            try:
                import sys
                sys.exit(0)
            except:
                pass
    
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
            if hasattr(self, 'input_text') and hasattr(self.input_text, '_textbox'):
                self.input_text._textbox.edit_undo()
                # 撤销后强制触发预览更新
                self.preview_sync.on_text_change(None)
                self.update_status("↶ 撤销成功")
        except tk.TclError:
            self.update_status("没有可撤销的操作")
        except Exception as e:
            print(f"Undo error: {e}")
    
    def _redo(self):
        """重做操作"""
        try:
            if hasattr(self, 'input_text') and hasattr(self.input_text, '_textbox'):
                self.input_text._textbox.edit_redo()
                # 重做后强制触发预览更新
                self.preview_sync.on_text_change(None)
                self.update_status("↷ 重做成功")
        except tk.TclError:
            self.update_status("没有可重做的操作")
        except Exception as e:
            print(f"Redo error: {e}")
    
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
        """显示插入菜单"""
        if hasattr(self, 'insert_templates'):
            menu = tk.Menu(self, tearoff=0)
            
            # 基础模板
            menu.add_command(label="📋 表格", command=self.insert_templates.insert_table_template)
            menu.add_command(label="🔗 链接", command=self.insert_templates.insert_link_template)
            menu.add_command(label="🖼️ 图片", command=self.insert_templates.insert_image_template)
            menu.add_command(label="∑ 公式", command=self.insert_templates.insert_math_template)
            menu.add_command(label="📝 代码块", command=self.insert_templates.insert_code_template)
            menu.add_command(label="☑️ 任务列表", command=self.insert_templates.insert_task_template)
            menu.add_command(label="➖ 分割线", command=self.insert_templates.insert_hr_template)
            
            # 分隔符
            menu.add_separator()
            
            # Phase 1 新增功能
            menu.add_command(label="📑 插入目录 (TOC)", command=self.toc_generator.show_toc_dialog)
            menu.add_command(label="💧 配置水印", command=self.watermark_feature.show_watermark_dialog)
            
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    def _create_header(self):
        """创建顶部标题栏"""
        build_header(self)

    def _legacy_create_main_content(self):
        """创建主内容区域"""
        build_main_content(self)

    def _legacy_create_sidebar_content(self):
        """创建侧边栏内容"""
        build_sidebar_content(self)

    def _legacy_create_input_panel(self, parent):
        """创建输入面板"""
        build_input_panel(self, parent)

    def _legacy_create_preview_panel(self, parent):
        """创建预览面板"""
        build_preview_panel(self, parent)

    
    

def main():
    """启动应用"""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
