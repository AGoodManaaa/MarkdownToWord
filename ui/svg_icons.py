# -*- coding: utf-8 -*-
"""SVG 图标加载器 - 支持 Tabler/Lucide 图标"""

import os
import io
from typing import Dict, Optional, Tuple
from PIL import Image, ImageTk

try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False


class SVGIconLoader:
    """SVG 图标加载器，支持颜色替换和缓存"""
    
    _cache: Dict[str, ImageTk.PhotoImage] = {}
    _icon_dir: str = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'assets', 'icons')
    
    @classmethod
    def set_icon_dir(cls, path: str):
        """设置图标目录"""
        cls._icon_dir = path
    
    @classmethod
    def load(cls, name: str, size: int = 24, color: str = "#FFFFFF") -> Optional[ImageTk.PhotoImage]:
        """加载 SVG 图标
        
        Args:
            name: 图标名称（不含扩展名）
            size: 图标尺寸
            color: 图标颜色（十六进制）
        
        Returns:
            PhotoImage 对象，加载失败返回 None
        """
        cache_key = f"{name}_{size}_{color}"
        if cache_key in cls._cache:
            return cls._cache[cache_key]
        
        if not HAS_CAIROSVG:
            return None
        
        svg_path = os.path.join(cls._icon_dir, f"{name}.svg")
        if not os.path.exists(svg_path):
            return None
        
        try:
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()
            
            # 替换颜色
            svg_content = svg_content.replace('currentColor', color)
            svg_content = svg_content.replace('#000000', color)
            svg_content = svg_content.replace('#000', color)
            
            # 转换为 PNG
            png_data = cairosvg.svg2png(
                bytestring=svg_content.encode('utf-8'),
                output_width=size,
                output_height=size
            )
            
            img = Image.open(io.BytesIO(png_data))
            photo = ImageTk.PhotoImage(img)
            cls._cache[cache_key] = photo
            return photo
        except Exception:
            return None
    
    @classmethod
    def clear_cache(cls):
        """清除图标缓存"""
        cls._cache.clear()


# Tabler 风格图标映射（名称 -> SVG 文件名）
TABLER_ICONS = {
    # 文件操作
    'file_new': 'file-plus',
    'file_open': 'folder-open',
    'file_save': 'device-floppy',
    'file_export': 'upload',
    'file_import': 'download',
    'file_pdf': 'file-type-pdf',
    'file_word': 'file-type-doc',
    
    # 编辑操作
    'edit_undo': 'arrow-back-up',
    'edit_redo': 'arrow-forward-up',
    'edit_cut': 'cut',
    'edit_copy': 'copy',
    'edit_paste': 'clipboard',
    'edit_delete': 'trash',
    'edit_find': 'search',
    'edit_replace': 'replace',
    
    # 视图
    'view_preview': 'eye',
    'view_sidebar': 'layout-sidebar-left',
    'view_minimap': 'map',
    'view_fullscreen': 'arrows-maximize',
    'view_split': 'layout-columns',
    'view_focus': 'focus-2',
    
    # 格式化
    'format_bold': 'bold',
    'format_italic': 'italic',
    'format_underline': 'underline',
    'format_strike': 'strikethrough',
    'format_code': 'code',
    'format_quote': 'quote',
    'format_link': 'link',
    'format_list_ul': 'list',
    'format_list_ol': 'list-numbers',
    'format_table': 'table',
    'format_heading': 'h-1',
    
    # 插入
    'insert_image': 'photo',
    'insert_table': 'table-plus',
    'insert_code': 'terminal-2',
    'insert_math': 'math-function',
    'insert_emoji': 'mood-smile',
    
    # 工具
    'tool_settings': 'settings',
    'tool_theme': 'palette',
    'tool_help': 'help-circle',
    'tool_info': 'info-circle',
    
    # AI
    'ai_assistant': 'robot',
    'ai_magic': 'sparkles',
    'ai_write': 'writing',
    
    # 主题
    'theme_light': 'sun',
    'theme_dark': 'moon',
    
    # 状态
    'status_success': 'circle-check',
    'status_error': 'circle-x',
    'status_warning': 'alert-triangle',
    'status_loading': 'loader-2',
    
    # 导航
    'nav_home': 'home',
    'nav_back': 'arrow-left',
    'nav_forward': 'arrow-right',
    'nav_folder': 'folder',
    'nav_file': 'file',
    
    # 其他
    'chart': 'chart-bar',
    'mindmap': 'sitemap',
    'bibliography': 'books',
    'version': 'git-branch',
    'database': 'database',
    'collab': 'users',
    'ocr': 'camera',
    'batch': 'stack-2',
}


def get_svg_icon(name: str, size: int = 24, color: str = "#FFFFFF") -> Optional[ImageTk.PhotoImage]:
    """获取 SVG 图标的便捷函数"""
    svg_name = TABLER_ICONS.get(name, name)
    return SVGIconLoader.load(svg_name, size, color)


# 导出
__all__ = ['SVGIconLoader', 'get_svg_icon', 'TABLER_ICONS', 'HAS_CAIROSVG']
