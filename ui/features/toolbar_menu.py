# -*- coding: utf-8 -*-
"""
工具栏下拉菜单组件
将多个工具按钮分组到下拉菜单中，减少视觉混乱
"""

import tkinter as tk
from typing import List, Tuple, Callable, Optional
import customtkinter as ctk

from ui.theme import COLORS
from ui.icons import get_toolbar_icon


class ToolbarDropdownMenu:
    """工具栏下拉菜单"""
    
    def __init__(self, parent, button_text: str, items: List[Tuple[str, str, Callable, str]], 
                 tooltip_manager=None, fg_color="transparent", hover_color=None):
        """
        初始化下拉菜单
        
        Args:
            parent: 父容器
            button_text: 按钮显示的文本/图标
            items: 菜单项列表 [(icon, label, command, shortcut), ...]
            tooltip_manager: 工具提示管理器
            fg_color: 按钮背景色
            hover_color: 悬停颜色
        """
        self.parent = parent
        self.items = items
        self.tooltip_manager = tooltip_manager
        self._menu_window = None
        
        # 创建触发按钮
        self.button = ctk.CTkButton(
            parent,
            text=button_text,
            width=50,
            height=34,
            corner_radius=10,
            fg_color=fg_color,
            hover_color=hover_color or COLORS.get('primary_hover', '#4F46E5'),
            text_color="white",
            font=("Segoe UI Emoji", 14),
            command=self._toggle_menu
        )
        
        # 添加下拉箭头指示
        self.button.configure(text=f"{button_text} ▾")
    
    def _toggle_menu(self):
        """切换菜单显示/隐藏"""
        if self._menu_window and self._menu_window.winfo_exists():
            self._close_menu()
        else:
            self._show_menu()
    
    def _show_menu(self):
        """显示下拉菜单"""
        if self._menu_window:
            self._close_menu()
        
        # 创建菜单窗口
        self._menu_window = tk.Toplevel(self.parent)
        self._menu_window.overrideredirect(True)  # 无边框
        self._menu_window.attributes('-topmost', True)
        
        # 计算位置
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height() + 2
        
        # 创建菜单框架
        menu_frame = ctk.CTkFrame(
            self._menu_window,
            fg_color=COLORS.get('bg_card', '#FFFFFF'),
            corner_radius=8,
            border_width=1,
            border_color=COLORS.get('border', '#E6E8F0')
        )
        menu_frame.pack(fill='both', expand=True, padx=1, pady=1)
        
        # 添加菜单项
        for icon, label, command, shortcut in self.items:
            item_frame = ctk.CTkFrame(menu_frame, fg_color="transparent")
            item_frame.pack(fill='x', padx=4, pady=2)
            
            # 菜单项按钮
            item_btn = ctk.CTkButton(
                item_frame,
                text=f"{icon}  {label}",
                anchor='w',
                width=180,
                height=32,
                corner_radius=6,
                fg_color="transparent",
                hover_color=COLORS.get('highlight', '#E0E7FF'),
                text_color=COLORS.get('text_primary', '#0F172A'),
                font=ctk.CTkFont(size=13),
                command=lambda cmd=command: self._on_item_click(cmd)
            )
            item_btn.pack(side='left', fill='x', expand=True)
            
            # 快捷键标签
            if shortcut:
                shortcut_label = ctk.CTkLabel(
                    item_frame,
                    text=shortcut,
                    text_color=COLORS.get('text_muted', '#94A3B8'),
                    font=ctk.CTkFont(size=11)
                )
                shortcut_label.pack(side='right', padx=8)
        
        # 设置窗口位置和大小
        self._menu_window.geometry(f"+{x}+{y}")
        
        # 点击其他地方关闭菜单
        self._menu_window.bind('<FocusOut>', lambda e: self._close_menu())
        self.parent.winfo_toplevel().bind('<Button-1>', self._on_click_outside, add='+')
    
    def _on_item_click(self, command: Callable):
        """菜单项点击"""
        self._close_menu()
        if command:
            command()
    
    def _on_click_outside(self, event):
        """点击菜单外部时关闭"""
        if self._menu_window and self._menu_window.winfo_exists():
            # 检查点击是否在菜单窗口内
            try:
                x, y = event.x_root, event.y_root
                menu_x = self._menu_window.winfo_rootx()
                menu_y = self._menu_window.winfo_rooty()
                menu_w = self._menu_window.winfo_width()
                menu_h = self._menu_window.winfo_height()
                
                if not (menu_x <= x <= menu_x + menu_w and menu_y <= y <= menu_y + menu_h):
                    # 也检查是否点击了触发按钮
                    btn_x = self.button.winfo_rootx()
                    btn_y = self.button.winfo_rooty()
                    btn_w = self.button.winfo_width()
                    btn_h = self.button.winfo_height()
                    
                    if not (btn_x <= x <= btn_x + btn_w and btn_y <= y <= btn_y + btn_h):
                        self._close_menu()
            except:
                pass
    
    def _close_menu(self):
        """关闭菜单"""
        if self._menu_window:
            try:
                self._menu_window.destroy()
            except:
                pass
            self._menu_window = None
    
    def pack(self, **kwargs):
        """Pack按钮"""
        self.button.pack(**kwargs)
    
    def grid(self, **kwargs):
        """Grid按钮"""
        self.button.grid(**kwargs)


class ToolbarSeparator(ctk.CTkFrame):
    """工具栏分隔线"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, width=1, height=24, fg_color=COLORS.get('border', '#E6E8F0'), **kwargs)


def create_grouped_toolbar(parent, app, tooltip_manager=None) -> dict:
    """
    创建分组的工具栏
    
    Args:
        parent: 父容器
        app: 应用实例
        tooltip_manager: 工具提示管理器
    
    Returns:
        包含所有按钮引用的字典
    """
    buttons = {}
    
    # 文件组 - 常用操作直接显示
    file_buttons = [
        (get_toolbar_icon("open"), "打开", app.open_file, "Ctrl+O"),
        (get_toolbar_icon("save"), "保存", app.save_file, "Ctrl+S"),
    ]
    
    for icon, tip, cmd, shortcut in file_buttons:
        btn = ctk.CTkButton(
            parent,
            text=icon,
            width=38,
            height=34,
            corner_radius=10,
            fg_color="transparent",
            hover_color=COLORS.get('primary_hover', '#4F46E5'),
            text_color="white",
            font=("Segoe UI Emoji", 16),
            command=cmd,
        )
        btn.pack(side="left", padx=2)
        if tooltip_manager:
            tooltip_manager.add_tooltip(btn, f"{tip}\n{shortcut}" if shortcut else tip)
        buttons[tip] = btn
    
    # 分隔线
    ToolbarSeparator(parent).pack(side="left", padx=6, pady=5)
    
    # 编辑组
    edit_buttons = [
        (get_toolbar_icon("format"), "规范化", app.format_markdown, "Ctrl+Shift+F"),
        (get_toolbar_icon("search"), "搜索", app.show_search_dialog, "Ctrl+F"),
    ]
    
    for icon, tip, cmd, shortcut in edit_buttons:
        btn = ctk.CTkButton(
            parent,
            text=icon,
            width=38,
            height=34,
            corner_radius=10,
            fg_color="transparent",
            hover_color=COLORS.get('primary_hover', '#4F46E5'),
            text_color="white",
            font=("Segoe UI Emoji", 16),
            command=cmd,
        )
        btn.pack(side="left", padx=2)
        if tooltip_manager:
            tooltip_manager.add_tooltip(btn, f"{tip}\n{shortcut}" if shortcut else tip)
        buttons[tip] = btn
    
    # 分隔线
    ToolbarSeparator(parent).pack(side="left", padx=6, pady=5)
    
    # 预览按钮
    preview_btn = ctk.CTkButton(
        parent,
        text=get_toolbar_icon("preview"),
        width=38,
        height=34,
        corner_radius=10,
        fg_color="transparent",
        hover_color=COLORS.get('primary_hover', '#4F46E5'),
        text_color="white",
        font=("Segoe UI Emoji", 16),
        command=app.toggle_preview,
    )
    preview_btn.pack(side="left", padx=2)
    if tooltip_manager:
        tooltip_manager.add_tooltip(preview_btn, "预览\nCtrl+P")
    buttons["预览"] = preview_btn
    
    # 分隔线
    ToolbarSeparator(parent).pack(side="left", padx=6, pady=5)
    
    # 导出下拉菜单
    export_items = [
        (get_toolbar_icon("export"), "导出Word", app.export_to_word, "Ctrl+Shift+S"),
        (get_toolbar_icon("pdf"), "导出PDF", app.export_to_pdf, ""),
        ("🌐", "导出HTML", app.show_html_export, ""),
        (get_toolbar_icon("batch"), "批量导出", app.show_batch_export, ""),
    ]
    export_menu = ToolbarDropdownMenu(
        parent, "📤", export_items, tooltip_manager,
        hover_color=COLORS.get('primary_hover', '#4F46E5')
    )
    export_menu.pack(side="left", padx=2)
    buttons["导出"] = export_menu
    
    # 工具下拉菜单
    tools_items = [
        (get_toolbar_icon("ocr"), "OCR识别", app.show_ocr, "Ctrl+Shift+O"),
        (get_toolbar_icon("ai"), "AI助手", app.show_ai_assistant, "Ctrl+I"),
        (get_toolbar_icon("chart"), "图表编辑", app.show_chart_editor, ""),
        (get_toolbar_icon("mindmap"), "思维导图", app.show_mindmap, ""),
        (get_toolbar_icon("bibliography"), "文献管理", app.show_bibliography, ""),
        (get_toolbar_icon("link"), "链接检查", app.show_link_checker, ""),
    ]
    tools_menu = ToolbarDropdownMenu(
        parent, "🛠️", tools_items, tooltip_manager,
        hover_color=COLORS.get('primary_hover', '#4F46E5')
    )
    tools_menu.pack(side="left", padx=2)
    buttons["工具"] = tools_menu
    
    # 协作下拉菜单
    collab_items = [
        (get_toolbar_icon("database"), "文档库", app.show_database, "Ctrl+Shift+D"),
        (get_toolbar_icon("collab"), "实时协作", app.show_collaboration, "Ctrl+Alt+C"),
        (get_toolbar_icon("version"), "版本控制", app.show_version_control, ""),
    ]
    collab_menu = ToolbarDropdownMenu(
        parent, "👥", collab_items, tooltip_manager,
        hover_color=COLORS.get('primary_hover', '#4F46E5')
    )
    collab_menu.pack(side="left", padx=2)
    buttons["协作"] = collab_menu
    
    return buttons
