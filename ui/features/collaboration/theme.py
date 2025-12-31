# -*- coding: utf-8 -*-
"""协作功能主题管理模块 - 支持深色/浅色模式"""

import json
import os
from typing import Dict, Optional, Callable
import darkdetect

# 主题配置文件路径
THEME_CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'theme_config.json')

# 浅色主题配色
LIGHT_THEME = {
    'name': 'light',
    'primary': '#10b981',
    'primary_dark': '#059669',
    'primary_light': '#d1fae5',
    'secondary': '#3b82f6',
    'secondary_dark': '#2563eb',
    'secondary_light': '#dbeafe',
    'danger': '#ef4444',
    'danger_dark': '#dc2626',
    'danger_light': '#fee2e2',
    'warning': '#f59e0b',
    'warning_light': '#fef3c7',
    'success': '#10b981',
    'success_light': '#d1fae5',
    'gray': '#6b7280',
    'gray_light': '#f3f4f6',
    'gray_dark': '#374151',
    'text': '#1f2937',
    'text_secondary': '#6b7280',
    'text_muted': '#9ca3af',
    'background': '#ffffff',
    'background_secondary': '#f9fafb',
    'surface': '#ffffff',
    'surface_hover': '#f3f4f6',
    'border': '#e5e7eb',
    'border_light': '#f3f4f6',
    'shadow': 'rgba(0,0,0,0.1)',
    'overlay': 'rgba(0,0,0,0.5)',
    'toast_info': ('#3b82f6', '#dbeafe'),
    'toast_success': ('#10b981', '#d1fae5'),
    'toast_warning': ('#f59e0b', '#fef3c7'),
    'toast_error': ('#ef4444', '#fee2e2'),
}

# 深色主题配色
DARK_THEME = {
    'name': 'dark',
    'primary': '#34d399',
    'primary_dark': '#10b981',
    'primary_light': '#064e3b',
    'secondary': '#60a5fa',
    'secondary_dark': '#3b82f6',
    'secondary_light': '#1e3a5f',
    'danger': '#f87171',
    'danger_dark': '#ef4444',
    'danger_light': '#7f1d1d',
    'warning': '#fbbf24',
    'warning_light': '#78350f',
    'success': '#34d399',
    'success_light': '#064e3b',
    'gray': '#9ca3af',
    'gray_light': '#374151',
    'gray_dark': '#1f2937',
    'text': '#f9fafb',
    'text_secondary': '#d1d5db',
    'text_muted': '#9ca3af',
    'background': '#111827',
    'background_secondary': '#1f2937',
    'surface': '#1f2937',
    'surface_hover': '#374151',
    'border': '#374151',
    'border_light': '#4b5563',
    'shadow': 'rgba(0,0,0,0.3)',
    'overlay': 'rgba(0,0,0,0.7)',
    'toast_info': ('#60a5fa', '#1e3a5f'),
    'toast_success': ('#34d399', '#064e3b'),
    'toast_warning': ('#fbbf24', '#78350f'),
    'toast_error': ('#f87171', '#7f1d1d'),
}


class ThemeManager:
    """主题管理器"""
    
    _instance: Optional['ThemeManager'] = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        
        self._current_theme = 'light'
        self._auto_mode = True
        self._listeners: list[Callable] = []
        self._load_config()
        
        # 如果是自动模式，检测系统主题
        if self._auto_mode:
            self._detect_system_theme()
    
    def _load_config(self):
        """加载主题配置"""
        try:
            if os.path.exists(THEME_CONFIG_FILE):
                with open(THEME_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self._current_theme = config.get('theme', 'light')
                    self._auto_mode = config.get('auto_mode', True)
        except Exception:
            pass
    
    def _save_config(self):
        """保存主题配置"""
        try:
            config = {
                'theme': self._current_theme,
                'auto_mode': self._auto_mode
            }
            with open(THEME_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    
    def _detect_system_theme(self):
        """检测系统主题"""
        try:
            system_theme = darkdetect.theme()
            if system_theme:
                self._current_theme = system_theme.lower()
        except Exception:
            pass
    
    @property
    def current_theme(self) -> str:
        """获取当前主题名称"""
        return self._current_theme
    
    @property
    def colors(self) -> Dict:
        """获取当前主题颜色"""
        return DARK_THEME if self._current_theme == 'dark' else LIGHT_THEME
    
    @property
    def is_dark(self) -> bool:
        """是否为深色模式"""
        return self._current_theme == 'dark'
    
    @property
    def auto_mode(self) -> bool:
        """是否为自动模式"""
        return self._auto_mode
    
    def set_theme(self, theme: str):
        """设置主题"""
        if theme in ('light', 'dark'):
            self._current_theme = theme
            self._auto_mode = False
            self._save_config()
            self._notify_listeners()
    
    def set_auto_mode(self, enabled: bool):
        """设置自动模式"""
        self._auto_mode = enabled
        if enabled:
            self._detect_system_theme()
        self._save_config()
        self._notify_listeners()
    
    def toggle_theme(self):
        """切换主题"""
        new_theme = 'light' if self._current_theme == 'dark' else 'dark'
        self.set_theme(new_theme)
    
    def add_listener(self, callback: Callable):
        """添加主题变化监听器"""
        if callback not in self._listeners:
            self._listeners.append(callback)
    
    def remove_listener(self, callback: Callable):
        """移除监听器"""
        if callback in self._listeners:
            self._listeners.remove(callback)
    
    def _notify_listeners(self):
        """通知所有监听器"""
        for listener in self._listeners:
            try:
                listener(self.colors)
            except Exception:
                pass


# 全局主题管理器实例
theme_manager = ThemeManager()


def get_colors() -> Dict:
    """获取当前主题颜色 - 固定使用浅色主题"""
    # 用户要求不使用深色模式，固定返回浅色主题
    return LIGHT_THEME


def is_dark_mode() -> bool:
    """是否为深色模式 - 固定返回 False"""
    return False
