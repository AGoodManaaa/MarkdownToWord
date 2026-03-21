# -*- coding: utf-8 -*-

import tkinter as tk
import customtkinter as ctk

from ui.theme import COLORS
from ui.widgets import ModernCard, ModernButton
from ui.sidebar import OutlineView, RecentFilesView
from ui.icons import get_toolbar_icon
from ui.editor import LineNumberedText
from ui.preview import MarkdownPreview
from ui.features.toolbar_menu import ToolbarDropdownMenu


def build_header(app) -> None:
    app.header = ctk.CTkFrame(app, fg_color=COLORS["primary"], height=60, corner_radius=0)
    app.header.pack(fill="x", side="top")
    app.header.pack_propagate(False)

    left_frame = ctk.CTkFrame(app.header, fg_color="transparent")
    left_frame.pack(side="left", padx=20, pady=12)

    app.title_label = ctk.CTkLabel(
        left_frame,
        text="Markdown -> Word",
        font=ctk.CTkFont(size=18, weight="bold"),
        text_color="white",
    )
    app.title_label.pack(side="left")

    toolbar_frame = ctk.CTkFrame(app.header, fg_color="transparent")
    toolbar_frame.pack(side="left", padx=24)

    tools = [
        (get_toolbar_icon("open"), "打开", app.open_file, "Ctrl+O"),
        (get_toolbar_icon("save"), "保存", app.save_file, "Ctrl+S"),
        (get_toolbar_icon("format"), "规范化", app.format_markdown, "Ctrl+Shift+F"),
        (get_toolbar_icon("search"), "搜索", app.show_search_dialog, "Ctrl+F"),
        (get_toolbar_icon("preview"), "预览", app.toggle_preview, "Ctrl+P"),
        (get_toolbar_icon("split"), "左右分屏", app.set_split_horizontal, ""),
        ("↕", "上下分屏", app.set_split_vertical, ""),
        (get_toolbar_icon("export"), "导出 Word", app.export_to_word, "Ctrl+Shift+S"),
        (get_toolbar_icon("pdf"), "导出 PDF", app.export_to_pdf, ""),
        (get_toolbar_icon("batch"), "批量导出", app.show_batch_export, ""),
    ]

    app.preview_btn = None
    for icon, tip, cmd, shortcut in tools:
        btn = ctk.CTkButton(
            toolbar_frame,
            text=icon,
            width=38,
            height=34,
            corner_radius=10,
            fg_color="transparent",
            hover_color=COLORS["primary_hover"],
            text_color="white",
            font=("Segoe UI Emoji", 16),
            command=cmd,
        )
        btn.pack(side="left", padx=2)
        app._header_default_buttons.append(btn)
        app.tooltip.add_tooltip(btn, f"{tip}\n{shortcut}" if shortcut else tip)
        if tip == "预览":
            app.preview_btn = btn

    menu_items = [
        (get_toolbar_icon("split"), "左右分屏", app.set_split_horizontal, ""),
        ("↕", "上下分屏", app.set_split_vertical, ""),
        ("E", "仅编辑", app.set_editor_only, ""),
        ("P", "仅预览", app.set_preview_only, ""),
        (get_toolbar_icon("fullscreen"), "全屏预览", app.toggle_fullscreen_preview, "Ctrl+F11"),
        ("⌨", "快捷键", app.show_keyboard_shortcuts, ""),
        (get_toolbar_icon("minimap"), "迷你地图", app.toggle_minimap, ""),
        (get_toolbar_icon("html"), "HTML 导出", app.show_html_export, ""),
        (get_toolbar_icon("ocr"), "OCR", app.show_ocr, "Ctrl+Shift+O"),
        (get_toolbar_icon("ai"), "AI 助手", app.show_ai_assistant, "Ctrl+I"),
    ]
    if getattr(app, "show_advanced_toolbar", False):
        menu_items.extend([
            (get_toolbar_icon("chart"), "图表", app.show_chart_editor, ""),
            (get_toolbar_icon("mindmap"), "导图", app.show_mindmap, ""),
            (get_toolbar_icon("bibliography"), "文献", app.show_bibliography, ""),
            (get_toolbar_icon("version"), "版本历史", app.show_version_control, ""),
            (get_toolbar_icon("link"), "链接检查", app.show_link_checker, ""),
            (get_toolbar_icon("database"), "文档库", app.show_database, "Ctrl+Shift+D"),
            (get_toolbar_icon("collab"), "协作", app.show_collaboration, "Ctrl+Alt+C"),
        ])

    app.advanced_tools_menu = ToolbarDropdownMenu(
        toolbar_frame,
        "高级",
        menu_items,
        tooltip_manager=app.tooltip,
        fg_color="transparent",
    )
    app.advanced_tools_menu.pack(side="left", padx=(8, 2))
    app._header_default_buttons.append(app.advanced_tools_menu.button)
    app.tooltip.add_tooltip(
        app.advanced_tools_menu.button,
        "高级工具\n收纳分屏、HTML、OCR、AI 等扩展能力",
    )

    app.insert_btn = ctk.CTkButton(
        toolbar_frame,
        text=get_toolbar_icon("insert"),
        width=38,
        height=34,
        corner_radius=10,
        fg_color=COLORS["success"],
        hover_color="#16A34A",
        text_color="white",
        font=ctk.CTkFont(size=15, weight="bold"),
        command=lambda: None,
    )
    app.insert_btn.pack(side="left", padx=2)
    app.insert_btn.bind("<Button-1>", app.show_insert_menu)
    app.tooltip.add_tooltip(app.insert_btn, "插入\n点击选择常用模板")

    btn_frame = ctk.CTkFrame(app.header, fg_color="transparent")
    btn_frame.pack(side="right", padx=20, pady=12)

    right_buttons = [
        ("sidebar_btn", get_toolbar_icon("sidebar"), app.toggle_sidebar, "侧边栏\nCtrl+B", 15),
        ("font_minus_btn", get_toolbar_icon("font_minus"), lambda: app.change_font_size(-1), "字号缩小\nCtrl+-", 12),
        ("font_plus_btn", get_toolbar_icon("font_plus"), lambda: app.change_font_size(1), "字号放大\nCtrl++", 12),
        ("focus_mode_btn", get_toolbar_icon("focus"), app.focus_mode.toggle, "专注模式\nF11", 15),
        ("reading_mode_btn", get_toolbar_icon("reading"), app.reading_mode.toggle, "阅读模式\nF12", 15),
        ("export_style_header_btn", get_toolbar_icon("settings"), app.open_export_style_settings, "导出样式设置\n含 Word 模板", 15),
    ]

    for attr_name, icon, cmd, tooltip_text, font_size in right_buttons:
        btn = ctk.CTkButton(
            btn_frame,
            text=icon,
            command=cmd,
            fg_color="transparent",
            text_color="white",
            hover_color=COLORS["primary_hover"],
            corner_radius=10,
            width=38,
            height=34,
            font=ctk.CTkFont(size=font_size, weight="bold"),
        )
        btn.pack(side="left", padx=3 if attr_name not in {"font_minus_btn", "font_plus_btn"} else 1)
        app._header_default_buttons.append(btn)
        app.tooltip.add_tooltip(btn, tooltip_text)
        setattr(app, attr_name, btn)

    app.header_styler.update_states()


def build_main_content(app) -> None:
    app.main_container = ctk.CTkFrame(app, fg_color="transparent")
    app.main_container.pack(fill="both", expand=True, padx=15, pady=15)

    app.sidebar_visible = app.config.get("sidebar_visible", True)
    app.sidebar = ctk.CTkFrame(app.main_container, fg_color=COLORS["bg_sidebar"], width=250, corner_radius=12)
    if app.sidebar_visible:
        app.sidebar.pack(side="left", fill="y", padx=(0, 10))
    app.sidebar.pack_propagate(False)

    build_sidebar_content(app)

    app.right_container = ctk.CTkFrame(app.main_container, fg_color="transparent")
    app.right_container.pack(side="left", fill="both", expand=True)

    app.tab_bar = app.tab_manager.create_tab_bar(app.right_container)
    app.tab_bar.pack(fill="x", pady=(0, 4))

    app.main_frame = ctk.CTkFrame(app.right_container, fg_color="transparent")
    app.main_frame.pack(fill="both", expand=True)

    app.paned_window = tk.PanedWindow(
        app.main_frame,
        orient=tk.HORIZONTAL,
        bg=COLORS["bg_light"],
        bd=0,
        sashwidth=4,
        sashrelief="flat",
    )
    app.paned_window.pack(fill="both", expand=True)
    app.paned_window.bind("<ButtonRelease-1>", app._on_split_drag_end)

    app._create_input_panel(app.paned_window)
    app.paned_window.add(app.input_card, stretch="always")

    app._create_preview_panel(app.paned_window)
    app.paned_window.add(app.preview_card, stretch="always")

    def _set_half_half():
        try:
            width = app.paned_window.winfo_width()
            if width <= 0:
                app.after(100, _set_half_half)
                return
            ratio = float(app.config.get("split_ratio", 0.5))
            app.paned_window.sash_place(0, int(width * ratio), 0)
        except Exception:
            pass

    app.after(150, _set_half_half)
    app._insert_example()


def build_sidebar_content(app) -> None:
    app.folder_view = None
    folder_feature = app._ensure_optional_feature("folder_view_feature", "文件夹视图")
    if folder_feature:
        try:
            app.folder_view = folder_feature.create_view(app.sidebar)
            app.folder_view.pack(fill="both", expand=True, pady=(0, 5))
        except Exception:
            app.folder_view = None

    if app.folder_view is None:
        folder_placeholder = ctk.CTkFrame(app.sidebar, fg_color="transparent")
        folder_placeholder.pack(fill="x", padx=12, pady=(0, 5))
        ctk.CTkLabel(
            folder_placeholder,
            text="文件夹视图未启用",
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(fill="x", pady=(6, 2))
        ctk.CTkButton(
            folder_placeholder,
            text="打开文件夹",
            height=30,
            command=app.open_folder,
        ).pack(fill="x")

    separator1 = ctk.CTkFrame(app.sidebar, height=1, fg_color=COLORS["border"])
    separator1.pack(fill="x", padx=15, pady=5)

    app.outline_view = OutlineView(app.sidebar, on_heading_click=app._jump_to_line)
    app.outline_view.pack(fill="both", expand=True, pady=(0, 10))

    separator2 = ctk.CTkFrame(app.sidebar, height=1, fg_color=COLORS["border"])
    separator2.pack(fill="x", padx=15, pady=5)

    app.recent_files_view = RecentFilesView(app.sidebar, on_file_click=app._open_recent_file)
    app.recent_files_view.pack(fill="both", expand=True)


def build_input_panel(app, parent) -> None:
    app.input_card = ModernCard(parent)

    toolbar = ctk.CTkFrame(app.input_card, fg_color="transparent", height=26)
    toolbar.pack(fill="x", padx=6, pady=(6, 0))
    toolbar.pack_propagate(False)

    editor_zoom_controls = app.editor_zoom_feature.create_controls(toolbar)
    editor_zoom_controls.pack(side="right", padx=(4, 0))

    groups = [
        [("H1", "# "), ("H2", "## "), ("H3", "### ")],
        [("B", "**bold**"), ("I", "*italic*"), ("~", "~~delete~~")],
        [("^", "<sup>sup</sup>"), ("_", "<sub>sub</sub>")],
        [("Img", "![image](url)"), ("Link", "[link](url)"), ("Math", "$formula$")],
        [("Tbl", "| header |\n|---|\n| content |"), ("Code", "```python\ncode\n```")],
    ]

    for index, group in enumerate(groups):
        if index > 0:
            sep = ctk.CTkFrame(toolbar, width=1, fg_color=COLORS["border"])
            sep.pack(side="left", fill="y", padx=3, pady=2)

        for text, insert_text in group:
            btn = ctk.CTkButton(
                toolbar,
                text=text,
                width=36 if len(text) > 1 else 26,
                height=22,
                corner_radius=8,
                fg_color=COLORS["bg_card"],
                text_color=COLORS["text_primary"],
                hover_color=COLORS["highlight"],
                border_width=1,
                border_color=COLORS["border"],
                font=ctk.CTkFont(size=10, weight="bold"),
                command=lambda t=insert_text: app.insert_text(t),
            )
            btn.pack(side="left", padx=1)

    app.welcome_panel = ctk.CTkFrame(
        app.input_card,
        fg_color=COLORS["bg_sidebar"],
        corner_radius=12,
        border_width=1,
        border_color=COLORS["border"],
    )
    app.welcome_panel.pack(fill="x", padx=6, pady=(8, 4))
    app.welcome_panel._visible = True

    welcome_header = ctk.CTkFrame(app.welcome_panel, fg_color="transparent")
    welcome_header.pack(fill="x", padx=16, pady=(14, 8))

    ctk.CTkLabel(
        welcome_header,
        text="开始一次高质量转换",
        font=ctk.CTkFont(size=20, weight="bold"),
        text_color=COLORS["text_primary"],
        anchor="w",
    ).pack(anchor="w")

    ctk.CTkLabel(
        welcome_header,
        text="导入 Markdown，检查预览，选择模板，然后直接导出 Word 或 PDF。",
        font=ctk.CTkFont(size=12),
        text_color=COLORS["text_secondary"],
        anchor="w",
    ).pack(anchor="w", pady=(4, 0))

    action_row = ctk.CTkFrame(app.welcome_panel, fg_color="transparent")
    action_row.pack(fill="x", padx=16, pady=(0, 10))

    ctk.CTkButton(
        action_row,
        text="打开 Markdown",
        command=app.open_file,
        fg_color=COLORS["primary"],
        hover_color=COLORS["primary_hover"],
        height=34,
        width=126,
    ).pack(side="left", padx=(0, 8))

    ctk.CTkButton(
        action_row,
        text="粘贴内容",
        command=app.paste_from_clipboard,
        fg_color=COLORS["bg_card"],
        text_color=COLORS["text_primary"],
        hover_color=COLORS["highlight"],
        border_width=1,
        border_color=COLORS["border"],
        height=34,
        width=110,
    ).pack(side="left", padx=(0, 8))

    ctk.CTkButton(
        action_row,
        text="插入示例",
        command=app.load_welcome_sample,
        fg_color=COLORS["bg_card"],
        text_color=COLORS["text_primary"],
        hover_color=COLORS["highlight"],
        border_width=1,
        border_color=COLORS["border"],
        height=34,
        width=110,
    ).pack(side="left")

    for item in [
        "1. 先检查标题层级、图片路径和表格列数",
        "2. 再看右侧预览，确认分页与样式",
        "3. 最后在样式设置里选模板后导出",
    ]:
        ctk.CTkLabel(
            app.welcome_panel,
            text=item,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(fill="x", padx=16, pady=1)

    app.input_editor = LineNumberedText(
        app.input_card,
        font_size=app.config.get("font_size", 14),
        on_scroll=app._on_editor_scroll,
    )
    app.input_editor.pack(fill="both", expand=True, padx=6, pady=(4, 6))
    app.input_text = app.input_editor

    if hasattr(app.input_editor, "text"):
        app.input_editor.text.bind("<KeyRelease>", app._on_text_change_debounced, add="+")
    else:
        app.input_editor.bind("<KeyRelease>", app._on_text_change_debounced)

    app.input_editor.bind("<Tab>", app.insert_templates.on_tab)
    app.input_editor.bind("<Escape>", app.insert_templates.on_escape)
    app.input_editor.bind("<KeyRelease>", app._on_cursor_event)
    app.input_editor.bind("<ButtonRelease-1>", app._on_cursor_event)

    try:
        app.editor_context_menu_feature.attach(app.input_editor._textbox)
    except Exception:
        pass

    try:
        app._update_welcome_state()
    except Exception:
        pass


def build_preview_panel(app, parent) -> None:
    app.preview_visible = True
    app.preview_card = ModernCard(parent, title="实时预览")

    top_frame = ctk.CTkFrame(app.preview_card, fg_color="transparent", height=30)
    top_frame.pack(fill="x", padx=10, pady=(4, 0))

    app.preview_zoom_controls = app.preview_zoom_feature.create_controls(top_frame)
    app.preview_zoom_controls.pack(side="right", padx=(0, 4))

    app.preview = MarkdownPreview(
        app.preview_card,
        on_content_change=app._on_preview_change,
        app=app,
        on_scroll=app._on_preview_scroll,
    )
    app.preview.pack(fill="both", expand=True, padx=10, pady=(6, 10))

    hint_frame = ctk.CTkFrame(app.preview_card, fg_color="transparent")
    hint_frame.pack(fill="x", padx=10, pady=(0, 6))
    hint_label = ctk.CTkLabel(
        hint_frame,
        text="预览区只读，支持 Ctrl+C，双击可跳转到源代码位置。",
        anchor="w",
        text_color=COLORS.get("text_secondary", "#6b7280"),
        font=ctk.CTkFont(size=11),
    )
    hint_label.pack(side="left", fill="x")
    app.preview.set_jump_callback(app._jump_to_line)

    btn_frame = ctk.CTkFrame(app.preview_card, fg_color="transparent", height=45)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))
    left_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
    left_group.pack(side="left")
    right_group = ctk.CTkFrame(btn_frame, fg_color="transparent")
    right_group.pack(side="right")

    app.export_btn = ModernButton(left_group, text="导出", command=app.export_to_word, style="primary", width=92)
    app.export_btn.pack(side="left", padx=(0, 8))

    app.cancel_export_btn = ModernButton(left_group, text="取消", command=app.cancel_export, style="ghost", width=86)
    app.cancel_export_btn.pack(side="left", padx=(0, 8))
    try:
        app.cancel_export_btn.configure(state="disabled")
        app.tooltip.add_tooltip(app.cancel_export_btn, "取消导出")
    except Exception:
        pass

    app.export_style_btn = ModernButton(left_group, text="样式", command=app.open_export_style_settings, style="ghost", width=52)
    app.export_style_btn.pack(side="left", padx=(0, 6))
    app.export_history_btn = ModernButton(left_group, text="历史", command=app.show_export_history, style="ghost", width=52)
    app.export_history_btn.pack(side="left", padx=(0, 6))
    app.copy_btn = ModernButton(left_group, text="复制", command=app.copy_to_clipboard, style="ghost", width=86)
    app.copy_btn.pack(side="left", padx=(0, 6))

    app.clear_btn = ModernButton(right_group, text="清空", command=app.clear_all, style="danger", width=60)
    app.clear_btn.pack(side="right", padx=(0, 8))
    app.refresh_preview_btn = ModernButton(right_group, text="刷新", command=app.refresh_preview, style="ghost", width=80)
    app.refresh_preview_btn.pack(side="right", padx=(0, 6))
