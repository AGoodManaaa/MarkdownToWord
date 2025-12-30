# -*- coding: utf-8 -*-
"""
对话框工具函数
提供统一的对话框设置功能，包括窗口图标设置等
"""

import os


# 获取 app.ico 路径
APP_ICO_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'app.ico')


def set_dialog_icon(dialog):
    """
    为 CTkToplevel 对话框设置应用图标
    
    Args:
        dialog: CTkToplevel 对话框实例
    """
    try:
        if os.path.exists(APP_ICO_PATH):
            # 在 Windows 上设置图标
            dialog.after(200, lambda: dialog.iconbitmap(APP_ICO_PATH))
    except Exception:
        pass  # 忽略图标设置失败


def center_dialog(dialog, parent, width=None, height=None):
    """
    将对话框居中显示在父窗口中心
    
    Args:
        dialog: CTkToplevel 对话框实例
        parent: 父窗口
        width: 对话框宽度（可选，如果对话框已设置尺寸则不需要）
        height: 对话框高度（可选，如果对话框已设置尺寸则不需要）
    """
    try:
        dialog.update_idletasks()
        
        if width is None:
            width = dialog.winfo_width()
        if height is None:
            height = dialog.winfo_height()
        
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - height) // 2
        
        dialog.geometry(f"+{x}+{y}")
    except Exception:
        pass


def setup_dialog(dialog, parent, title="", width=None, height=None):
    """
    统一设置对话框（图标、居中、标题等）
    
    Args:
        dialog: CTkToplevel 对话框实例
        parent: 父窗口
        title: 窗口标题
        width: 对话框宽度
        height: 对话框高度
    """
    if title:
        dialog.title(title)
    
    if width and height:
        dialog.geometry(f"{width}x{height}")
    
    dialog.transient(parent)
    
    # 设置图标
    set_dialog_icon(dialog)
    
    # 居中
    center_dialog(dialog, parent, width, height)
