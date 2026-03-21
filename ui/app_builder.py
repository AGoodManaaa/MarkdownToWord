# -*- coding: utf-8 -*-
"""
UI Builder - 统一 UI 构建入口

将 UI 构建逻辑从 App 类中分离出来，提高可维护性。
"""

import customtkinter as ctk
from ui.theme import COLORS
from ui.widgets import ModernButton, ModernCard
from ui.editor import LineNumberedText
from ui.preview import MarkdownPreview
from ui.sidebar import OutlineView, RecentFilesView


class UIBuilder:
    """统一 UI 构建入口。"""
    
    def __init__(self, app):
        """
        初始化 UIBuilder。
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self.header_builder = HeaderBuilder(app)
        self.main_content_builder = MainContentBuilder(app)
    
    def build(self):
        """构建完整 UI。"""
        self.header_builder.build()
        self.app.status_bar_feature.create()
        self.main_content_builder.build()


class HeaderBuilder:
    """Header 构建器。"""
    
    def __init__(self, app):
        """
        初始化 HeaderBuilder。
        
        Args:
            app: 主应用实例
        """
        self.app = app
    
    def build(self):
        """构建 Header。"""
        self._create_frame()
        self._create_logo()
        self._create_toolbar()
        self._create_right_buttons()
    
    def _create_frame(self):
        """创建 Header 框架。"""
        self.app.header = ctk.CTkFrame(
            self.app, 
            fg_color=COLORS['primary'], 
            height=60, 
            corner_radius=0
        )
        self.app.header.pack(fill="x", side="top")
        self.app.header.pack_propagate(False)
    
    def _create_logo(self):
        """创建 Logo 和标题。"""
        left_frame = ctk.CTkFrame(self.app.header, fg_color="transparent")
        left_frame.pack(side="left", padx=20, pady=12)
        
        self.app.title_label = ctk.CTkLabel(
            left_frame,
            text="📝 Markdown → Word",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        )
        self.app.title_label.pack(side="left")
    
    def _create_toolbar(self):
        """创建工具栏按钮。"""
        toolbar_frame = ctk.CTkFrame(self.app.header, fg_color="transparent")
        toolbar_frame.pack(side="left", padx=24)
        
        tools = self._get_tool_definitions()
        
        self.app.preview_btn = None
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
            self.app._header_default_buttons.append(btn)
            self.app.tooltip.add_tooltip(btn, f"{tip}\n{shortcut}")
            if tip == "预览":
                self.app.preview_btn = btn
        
        # 插入按钮
        self._create_insert_button(toolbar_frame)
    
    def _create_insert_button(self, parent):
        """创建插入按钮。"""
        self.app.insert_btn = ctk.CTkButton(
            parent,
            text="➕",
            width=38,
            height=34,
            corner_radius=10,
            fg_color=COLORS['success'],
            hover_color="#16A34A",
            text_color="white",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda: None
        )
        self.app.insert_btn.pack(side="left", padx=2)
        self.app.insert_btn.bind('<Button-1>', self.app.show_insert_menu)
        self.app.tooltip.add_tooltip(self.app.insert_btn, "插入\n点击选择模板")
    
    def _create_right_buttons(self):
        """创建右侧按钮组。"""
        btn_frame = ctk.CTkFrame(self.app.header, fg_color="transparent")
        btn_frame.pack(side="right", padx=20, pady=12)
        
        buttons = self._get_right_button_definitions()
        
        for name, icon, cmd, tooltip_text in buttons:
            btn = ctk.CTkButton(
                btn_frame,
                text=icon,
                command=cmd,
                fg_color="transparent",
                text_color="white",
                hover_color=COLORS['primary_hover'],
                corner_radius=10,
                width=38,
                height=34,
                font=ctk.CTkFont(size=15 if len(icon) == 1 else 12, weight="bold")
            )
            btn.pack(side="left", padx=3)
            self.app._header_default_buttons.append(btn)
            self.app.tooltip.add_tooltip(btn, tooltip_text)
            setattr(self.app, f"{name}_btn", btn)
    
    def _get_tool_definitions(self):
        """获取工具按钮定义。"""
        tools = [
            ("📂", "打开", self.app.open_file, "Ctrl+O"),
            ("💾", "保存", self.app.save_file, "Ctrl+S"),
            ("✦", "规范化", self.app.format_markdown, "Ctrl+Shift+F"),
            ("🔍", "搜索", self.app.show_search_dialog, "Ctrl+F"),
            ("👁", "预览", self.app.toggle_preview, "Ctrl+P"),
            ("📤", "导出", self.app.export_to_word, "Ctrl+Shift+S"),
            ("📄", "PDF", self.app.export_to_pdf, ""),
            ("📦", "批量导出", self.app.show_batch_export, "Ctrl+B"),
        ]
        if getattr(self.app, 'show_advanced_toolbar', False):
            tools.extend([
                ("🤖", "AI助手", self.app.show_ai_assistant, "Ctrl+I"),
                ("📊", "图表", self.app.show_chart_editor, "Ctrl+G"),
                ("🧠", "导图", self.app.show_mindmap, "Ctrl+T"),
                ("📑", "文献", self.app.show_bibliography, "Ctrl+R"),
                ("🔄", "版本", self.app.show_version_control, "Ctrl+H"),
                ("🔗", "链接", self.app.show_link_checker, "Ctrl+L"),
            ])
        return tools
    
    def _get_right_button_definitions(self):
        """获取右侧按钮定义。"""
        return [
            ("sidebar", "☰", self.app.toggle_sidebar, "侧边栏\nCtrl+B"),
            ("font_minus", "A-", lambda: self.app.change_font_size(-1), "字体大小\nCtrl++ / Ctrl+-"),
            ("font_plus", "A+", lambda: self.app.change_font_size(1), "字体大小\nCtrl++ / Ctrl+-"),
            ("theme", "🌙", self.app.toggle_theme, "切换亮/暗主题"),
            ("theme_editor", "🎨", self.app.theme_editor.show_editor, "自定义主题编辑器"),
            ("focus_mode", "🎯", self.app.focus_mode.toggle, "专注模式\nF11"),
            ("reading_mode", "📖", self.app.reading_mode.toggle, "阅读模式\nF12"),
            ("export_style_header", "⚙", self.app.open_export_style_settings, "导出样式设置\n(含导入Word模板)"),
        ]


class MainContentBuilder:
    """主内容区域构建器。"""
    
    def __init__(self, app):
        """
        初始化 MainContentBuilder。
        
        Args:
            app: 主应用实例
        """
        self.app = app
    
    def build(self):
        """构建主内容区域。"""
        self._create_main_container()
        self._create_sidebar()
        self._create_right_container()
        self._insert_example()
    
    def _create_main_container(self):
        """创建主容器。"""
        self.app.main_container = ctk.CTkFrame(self.app, fg_color="transparent")
        self.app.main_container.pack(fill="both", expand=True, padx=15, pady=15)
    
    def _create_sidebar(self):
        """创建侧边栏。"""
        self.app.sidebar_visible = self.app.config.get('sidebar_visible', True)
        self.app.sidebar = ctk.CTkFrame(
            self.app.main_container, 
            fg_color=COLORS['bg_sidebar'], 
            width=250, 
            corner_radius=12
        )
        if self.app.sidebar_visible:
            self.app.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.app.sidebar.pack_propagate(False)
        
        self._create_sidebar_content()
    
    def _create_sidebar_content(self):
        """创建侧边栏内容。"""
        # 大纲视图
        self.app.outline_view = OutlineView(
            self.app.sidebar,
            on_heading_click=self.app._jump_to_line
        )
        self.app.outline_view.pack(fill="both", expand=True, pady=(0, 10))
        
        # 分隔线
        separator = ctk.CTkFrame(self.app.sidebar, height=1, fg_color=COLORS['border'])
        separator.pack(fill="x", padx=15, pady=5)
        
        # 最近文件
        self.app.recent_files_view = RecentFilesView(
            self.app.sidebar,
            on_file_click=self.app._open_recent_file
        )
        self.app.recent_files_view.pack(fill="both", expand=True)
    
    def _create_right_container(self):
        """创建右侧主编辑区。"""
        self.app.right_container = ctk.CTkFrame(self.app.main_container, fg_color="transparent")
        self.app.right_container.pack(side="left", fill="both", expand=True)
        
        # 标签栏
        self.app.tab_bar = self.app.tab_manager.create_tab_bar(self.app.right_container)
        self.app.tab_bar.pack(fill="x", pady=(0, 4))
        
        # 编辑/预览区
        self.app.main_frame = ctk.CTkFrame(self.app.right_container, fg_color="transparent")
        self.app.main_frame.pack(fill="both", expand=True)
        
        # 配置列权重
        self.app.main_frame.grid_columnconfigure(0, weight=3)
        self.app.main_frame.grid_columnconfigure(1, weight=2)
        self.app.main_frame.grid_rowconfigure(0, weight=1)
        
        self._create_input_panel()
        self._create_preview_panel()
    
    def _create_input_panel(self):
        """创建输入面板。"""
        self.app.input_card = ModernCard(self.app.main_frame)
        self.app.input_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        
        # 工具栏
        toolbar = ctk.CTkFrame(self.app.input_card, fg_color="transparent", height=26)
        toolbar.pack(fill="x", padx=6, pady=(6, 0))
        toolbar.pack_propagate(False)
        
        # 编辑区缩放控件
        editor_zoom_controls = self.app.editor_zoom_feature.create_controls(toolbar)
        editor_zoom_controls.pack(side="right", padx=(4, 0))
        
        # 快捷插入按钮
        self._create_quick_insert_buttons(toolbar)
        
        # 带行号的输入文本框
        self.app.input_editor = LineNumberedText(
            self.app.input_card,
            font_size=self.app.config.get('font_size', 14),
            on_scroll=self.app._on_editor_scroll
        )
        self.app.input_editor.pack(fill="both", expand=True, padx=6, pady=(4, 6))
        
        # 兼容旧属性名
        self.app.input_text = self.app.input_editor
        
        # 绑定事件
        self.app.input_editor.bind('<KeyRelease>', self.app._on_text_change_debounced)
        self.app.input_editor.bind('<Tab>', self.app.insert_templates.on_tab)
        self.app.input_editor.bind('<Escape>', self.app.insert_templates.on_escape)
        self.app.input_editor.bind('<KeyRelease>', self.app._on_cursor_event)
        self.app.input_editor.bind('<ButtonRelease-1>', self.app._on_cursor_event)
        
        # 编辑器右键菜单
        try:
            self.app.editor_context_menu_feature.attach(self.app.input_editor._textbox)
        except Exception:
            pass
    
    def _create_quick_insert_buttons(self, toolbar):
        """创建快捷插入按钮。"""
        groups = [
            [("H1", "# "), ("H2", "## "), ("H3", "### ")],
            [("B", "**粗体**"), ("I", "*斜体*"), ("~", "~~删除~~")],
            [("²", "<sup>上标</sup>"), ("₂", "<sub>下标</sub>")],
            [("🖼", "![图片](url)"), ("🔗", "[链接](url)"), ("∑", "$公式$")],
            [("≣", "| 表头 |\n|---|\n| 内容 |"), ("`", "```python\ncode\n```")],
        ]
        
        for i, group in enumerate(groups):
            if i > 0:
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
                    command=lambda t=insert_text: self.app.insert_text(t)
                )
                btn.pack(side="left", padx=1)
    
    def _create_preview_panel(self):
        """创建预览面板。"""
        self.app.preview_visible = True
        self.app.preview_card = ModernCard(self.app.main_frame, title="👁️ 实时预览")
        self.app.preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        
        # 顶部缩放控件
        zoom_frame = ctk.CTkFrame(self.app.preview_card, fg_color="transparent", height=30)
        zoom_frame.pack(fill="x", padx=10, pady=(4, 0))
        
        zoom_controls = self.app.preview_zoom_feature.create_controls(zoom_frame)
        zoom_controls.pack(side="right")
        
        # 预览组件
        self.app.preview = MarkdownPreview(
            self.app.preview_card,
            on_content_change=self.app._on_preview_change,
            app=self.app,
            on_scroll=self.app._on_preview_scroll
        )
        self.app.preview.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        
        # 底部操作按钮
        self._create_preview_buttons()
    
    def _create_preview_buttons(self):
        """创建预览面板底部按钮。"""
        btn_frame = ctk.CTkFrame(self.app.preview_card, fg_color="transparent", height=45)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        left_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
        left_group.pack(side="left")
        
        right_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
        right_group.pack(side="right")
        
        # 导出按钮
        self.app.export_btn = ModernButton(
            left_group, text="📤 导出", command=self.app.export_to_word,
            style="primary", width=92
        )
        self.app.export_btn.pack(side="left", padx=(0, 8))
        
        # 取消导出按钮
        self.app.cancel_export_btn = ModernButton(
            left_group, text="⛔ 取消", command=self.app.cancel_export,
            style="ghost", width=86
        )
        self.app.cancel_export_btn.pack(side="left", padx=(0, 8))
        try:
            self.app.cancel_export_btn.configure(state="disabled")
            self.app.tooltip.add_tooltip(self.app.cancel_export_btn, "取消导出\n导出进行中可用")
        except Exception:
            pass
        
        # 其他按钮
        self.app.export_style_btn = ModernButton(
            left_group, text="⚙", command=self.app.open_export_style_settings,
            style="ghost", width=36
        )
        self.app.export_style_btn.pack(side="left", padx=(0, 6))
        
        self.app.export_history_btn = ModernButton(
            left_group, text="🕘", command=self.app.show_export_history,
            style="ghost", width=36
        )
        self.app.export_history_btn.pack(side="left", padx=(0, 6))
        
        self.app.copy_btn = ModernButton(
            left_group, text="📋 复制", command=self.app.copy_to_clipboard,
            style="ghost", width=86
        )
        self.app.copy_btn.pack(side="left", padx=(0, 6))
        
        self.app.clear_btn = ModernButton(
            right_group, text="🗑️", command=self.app.clear_all,
            style="danger", width=36
        )
        self.app.clear_btn.pack(side="right", padx=(0, 8))
    
    def _insert_example(self):
        """插入示例文本。"""
        from ui.startup_content import insert_example_if_empty_for_app
        insert_example_if_empty_for_app(self.app)
