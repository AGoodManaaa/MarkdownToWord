# -*- coding: utf-8 -*-
"""
Keyboard Shortcuts Manager - 快捷键统一管理

集中管理所有快捷键绑定，支持配置自定义。
"""

import logging
from typing import Any, Callable, Dict, Optional

logger = logging.getLogger(__name__)


class KeyboardShortcutsManager:
    """
    快捷键统一管理器。
    
    集中管理所有快捷键绑定，支持从配置加载自定义快捷键。
    """
    
    DEFAULT_SHORTCUTS = {
        '<Control-o>': 'open_file',
        '<Control-s>': 'save_file',
        '<Control-Shift-s>': 'export_to_word',
        '<Control-Shift-f>': 'format_markdown',
        '<Control-Shift-c>': 'copy_to_clipboard',
        '<Control-j>': 'show_export_history',
        '<Control-f>': 'show_search_dialog',
        '<Control-h>': 'show_search_dialog',
        '<Control-plus>': 'change_font_size:1',
        '<Control-minus>': 'change_font_size:-1',
        '<Control-b>': 'toggle_sidebar',
        '<Control-p>': 'toggle_preview',
        '<Control-z>': '_undo',
        '<Control-y>': '_redo',
        '<Control-Shift-z>': '_redo',
        '<Control-k>': 'command_palette.show',
        '<F1>': 'show_help',
        '<F11>': 'focus_mode.toggle',
        '<F12>': 'reading_mode.toggle',
        '<Control-F11>': 'toggle_fullscreen_preview',
        '<Control-Shift-p>': 'show_print_preview',
        '<Control-Shift-o>': 'show_ocr',
        '<Control-Shift-d>': 'show_database',
        '<Control-Alt-c>': 'show_collaboration',  # 协作功能使用 Ctrl+Alt+C 避免与复制冲突
    }
    
    def __init__(self, app):
        """
        初始化快捷键管理器。
        
        Args:
            app: 主应用实例
        """
        self.app = app
        self.shortcuts: Dict[str, str] = {}
        self._bound_handlers: Dict[str, Callable] = {}
    
    def load_from_config(self, config: Optional[Dict] = None) -> None:
        """
        从配置加载快捷键（支持自定义覆盖）。
        
        Args:
            config: 配置字典，可包含 'shortcuts' 键
        """
        self.shortcuts = self.DEFAULT_SHORTCUTS.copy()
        if config and 'shortcuts' in config:
            self.shortcuts.update(config['shortcuts'])
    
    def bind_all(self) -> None:
        """绑定所有快捷键到应用。"""
        for key, action in self.shortcuts.items():
            self._bind_shortcut(key, action)
    
    def unbind_all(self) -> None:
        """解绑所有快捷键。"""
        for key in self._bound_handlers:
            try:
                self.app.unbind(key)
            except Exception:
                pass
        self._bound_handlers.clear()
    
    def _bind_shortcut(self, key: str, action: str) -> bool:
        """
        绑定单个快捷键。
        
        Args:
            key: 快捷键字符串，如 '<Control-s>'
            action: 动作字符串，如 'save_file' 或 'focus_mode.toggle'
            
        Returns:
            True 如果绑定成功，否则 False
        """
        handler = self._resolve_action(action)
        if handler:
            try:
                self.app.bind(key, lambda e, h=handler: h())
                self._bound_handlers[key] = handler
                return True
            except Exception as e:
                logger.error(f"Failed to bind shortcut '{key}': {e}")
        return False
    
    def _resolve_action(self, action: str) -> Optional[Callable]:
        """
        解析 action 字符串为可调用对象。
        
        支持格式:
        - 'method_name': 调用 app.method_name()
        - 'feature.method': 调用 app.feature.method()
        - 'method:arg': 调用 app.method(arg)
        
        Args:
            action: 动作字符串
            
        Returns:
            可调用对象，如果解析失败则返回 None
        """
        # 处理带参数的动作
        if ':' in action:
            method_part, arg = action.split(':', 1)
            base_handler = self._resolve_action(method_part)
            if base_handler:
                try:
                    # 尝试转换参数为整数
                    arg_value = int(arg)
                    return lambda h=base_handler, a=arg_value: h(a)
                except ValueError:
                    return lambda h=base_handler, a=arg: h(a)
            return None
        
        # 处理嵌套属性访问
        parts = action.split('.')
        obj = self.app
        
        for part in parts:
            obj = getattr(obj, part, None)
            if obj is None:
                logger.warning(f"Cannot resolve action '{action}': '{part}' not found")
                return None
        
        return obj if callable(obj) else None
    
    def get_shortcut_for_action(self, action: str) -> Optional[str]:
        """
        获取指定动作的快捷键。
        
        Args:
            action: 动作字符串
            
        Returns:
            快捷键字符串，如果不存在则返回 None
        """
        for key, act in self.shortcuts.items():
            if act == action:
                return key
        return None
    
    def get_all_shortcuts(self) -> Dict[str, str]:
        """
        获取所有快捷键配置。
        
        Returns:
            快捷键配置字典的副本
        """
        return self.shortcuts.copy()
