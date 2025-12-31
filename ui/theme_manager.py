# -*- coding: utf-8 -*-
"""主题管理模块 - 加载和切换主题"""

import os
import json
import customtkinter as ctk
from typing import Optional

# 主题目录
THEMES_DIR = os.path.join(os.path.dirname(__file__), 'themes')


def get_available_themes() -> list:
    """获取所有可用的主题文件"""
    themes = []
    if os.path.exists(THEMES_DIR):
        for f in os.listdir(THEMES_DIR):
            if f.endswith('.json'):
                themes.append(f.replace('.json', ''))
    return themes


def load_theme(theme_name: str) -> bool:
    """加载指定主题
    
    Args:
        theme_name: 主题名称（不含扩展名）
    
    Returns:
        是否加载成功
    """
    theme_path = os.path.join(THEMES_DIR, f"{theme_name}.json")
    
    if not os.path.exists(theme_path):
        return False
    
    try:
        ctk.set_default_color_theme(theme_path)
        return True
    except Exception:
        return False


def apply_theme_to_app(app, theme_name: str = 'premium_dark'):
    """应用主题到整个应用
    
    Args:
        app: 应用实例
        theme_name: 主题名称
    """
    # 先加载主题
    load_theme(theme_name)
    
    # 更新 COLORS 字典
    from ui.theme import COLORS, COLORS_DARK, COLORS_LIGHT
    
    # 根据外观模式选择颜色
    mode = ctk.get_appearance_mode()
    if mode == "Dark":
        COLORS.clear()
        COLORS.update(COLORS_DARK)
    else:
        COLORS.clear()
        COLORS.update(COLORS_LIGHT)


def create_theme_preview(parent) -> ctk.CTkFrame:
    """创建主题预览卡片"""
    from ui.theme import COLORS
    
    frame = ctk.CTkFrame(parent, corner_radius=12)
    
    # 标题
    title = ctk.CTkLabel(
        frame, 
        text="主题预览",
        font=("微软雅黑", 16, "bold")
    )
    title.pack(pady=(15, 10))
    
    # 颜色预览
    colors_frame = ctk.CTkFrame(frame, fg_color="transparent")
    colors_frame.pack(fill="x", padx=15, pady=5)
    
    color_samples = [
        ('primary', COLORS.get('primary', '#818CF8')),
        ('success', COLORS.get('success', '#34D399')),
        ('warning', COLORS.get('warning', '#FBBF24')),
        ('danger', COLORS.get('danger', '#F87171')),
    ]
    
    for name, color in color_samples:
        color_box = ctk.CTkFrame(
            colors_frame,
            width=40,
            height=40,
            corner_radius=8,
            fg_color=color
        )
        color_box.pack(side="left", padx=5)
    
    # 按钮预览
    btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
    btn_frame.pack(fill="x", padx=15, pady=10)
    
    ctk.CTkButton(btn_frame, text="主按钮", width=100).pack(side="left", padx=5)
    ctk.CTkButton(
        btn_frame, 
        text="次按钮", 
        width=100,
        fg_color="transparent",
        border_width=1
    ).pack(side="left", padx=5)
    
    return frame


# 导出
__all__ = ['get_available_themes', 'load_theme', 'apply_theme_to_app', 'create_theme_preview']
