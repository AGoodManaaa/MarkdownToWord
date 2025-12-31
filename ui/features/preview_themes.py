# -*- coding: utf-8 -*-
"""预览主题模块 - 多种预览样式"""

from dataclasses import dataclass
from typing import Dict, Optional
import json
import os

# 主题配置文件路径
THEMES_DIR = os.path.join(os.path.dirname(__file__), 'preview_themes')


@dataclass
class PreviewTheme:
    """预览主题"""
    name: str
    display_name: str
    
    # 基础样式
    background: str = "#ffffff"
    text_color: str = "#1f2937"
    font_family: str = "Microsoft YaHei"
    font_size: int = 14
    line_height: float = 1.6
    
    # 标题样式
    h1_color: str = "#1f2937"
    h1_size: int = 28
    h1_weight: str = "bold"
    h1_border_bottom: str = "2px solid #e5e7eb"
    
    h2_color: str = "#374151"
    h2_size: int = 24
    h2_weight: str = "bold"
    h2_border_bottom: str = "1px solid #e5e7eb"
    
    h3_color: str = "#4b5563"
    h3_size: int = 20
    h3_weight: str = "bold"
    
    h4_color: str = "#6b7280"
    h4_size: int = 18
    h4_weight: str = "bold"
    
    h5_color: str = "#6b7280"
    h5_size: int = 16
    h5_weight: str = "bold"
    
    h6_color: str = "#9ca3af"
    h6_size: int = 14
    h6_weight: str = "bold"
    
    # 链接样式
    link_color: str = "#3b82f6"
    link_hover_color: str = "#2563eb"
    link_decoration: str = "none"
    
    # 代码样式
    code_bg: str = "#f3f4f6"
    code_color: str = "#dc2626"
    code_font: str = "Consolas"
    code_size: int = 13
    code_border_radius: str = "4px"
    
    # 代码块样式
    code_block_bg: str = "#1f2937"
    code_block_color: str = "#e5e7eb"
    code_block_border_radius: str = "8px"
    code_block_padding: str = "16px"
    
    # 引用样式
    blockquote_bg: str = "#f9fafb"
    blockquote_border: str = "4px solid #10b981"
    blockquote_color: str = "#6b7280"
    blockquote_padding: str = "12px 20px"
    
    # 表格样式
    table_border: str = "1px solid #e5e7eb"
    table_header_bg: str = "#f9fafb"
    table_header_color: str = "#374151"
    table_cell_padding: str = "12px 16px"
    table_stripe_bg: str = "#fafafa"
    
    # 列表样式
    list_marker_color: str = "#6b7280"
    list_indent: str = "24px"
    
    # 分隔线样式
    hr_color: str = "#e5e7eb"
    hr_height: str = "2px"
    
    # 图片样式
    image_border_radius: str = "8px"
    image_shadow: str = "0 4px 6px rgba(0,0,0,0.1)"
    
    def to_css(self) -> str:
        """生成 CSS 样式"""
        return f"""
        body {{
            background-color: {self.background};
            color: {self.text_color};
            font-family: '{self.font_family}', sans-serif;
            font-size: {self.font_size}px;
            line-height: {self.line_height};
            padding: 20px 40px;
            max-width: 900px;
            margin: 0 auto;
        }}
        
        h1 {{
            color: {self.h1_color};
            font-size: {self.h1_size}px;
            font-weight: {self.h1_weight};
            border-bottom: {self.h1_border_bottom};
            padding-bottom: 10px;
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        h2 {{
            color: {self.h2_color};
            font-size: {self.h2_size}px;
            font-weight: {self.h2_weight};
            border-bottom: {self.h2_border_bottom};
            padding-bottom: 8px;
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        h3 {{
            color: {self.h3_color};
            font-size: {self.h3_size}px;
            font-weight: {self.h3_weight};
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        h4 {{
            color: {self.h4_color};
            font-size: {self.h4_size}px;
            font-weight: {self.h4_weight};
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        h5 {{
            color: {self.h5_color};
            font-size: {self.h5_size}px;
            font-weight: {self.h5_weight};
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        h6 {{
            color: {self.h6_color};
            font-size: {self.h6_size}px;
            font-weight: {self.h6_weight};
            margin-top: 24px;
            margin-bottom: 16px;
        }}
        
        a {{
            color: {self.link_color};
            text-decoration: {self.link_decoration};
        }}
        
        a:hover {{
            color: {self.link_hover_color};
            text-decoration: underline;
        }}
        
        code {{
            background-color: {self.code_bg};
            color: {self.code_color};
            font-family: '{self.code_font}', monospace;
            font-size: {self.code_size}px;
            padding: 2px 6px;
            border-radius: {self.code_border_radius};
        }}
        
        pre {{
            background-color: {self.code_block_bg};
            color: {self.code_block_color};
            border-radius: {self.code_block_border_radius};
            padding: {self.code_block_padding};
            overflow-x: auto;
        }}
        
        pre code {{
            background-color: transparent;
            color: inherit;
            padding: 0;
        }}
        
        blockquote {{
            background-color: {self.blockquote_bg};
            border-left: {self.blockquote_border};
            color: {self.blockquote_color};
            padding: {self.blockquote_padding};
            margin: 16px 0;
        }}
        
        blockquote p {{
            margin: 0;
        }}
        
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 16px 0;
        }}
        
        th, td {{
            border: {self.table_border};
            padding: {self.table_cell_padding};
            text-align: left;
        }}
        
        th {{
            background-color: {self.table_header_bg};
            color: {self.table_header_color};
            font-weight: bold;
        }}
        
        tr:nth-child(even) {{
            background-color: {self.table_stripe_bg};
        }}
        
        ul, ol {{
            padding-left: {self.list_indent};
        }}
        
        li {{
            margin: 8px 0;
        }}
        
        li::marker {{
            color: {self.list_marker_color};
        }}
        
        hr {{
            border: none;
            height: {self.hr_height};
            background-color: {self.hr_color};
            margin: 24px 0;
        }}
        
        img {{
            max-width: 100%;
            border-radius: {self.image_border_radius};
            box-shadow: {self.image_shadow};
        }}
        
        /* 任务列表 */
        .task-list-item {{
            list-style-type: none;
            margin-left: -24px;
        }}
        
        .task-list-item input {{
            margin-right: 8px;
        }}
        """
    
    def to_tkinter_config(self) -> Dict:
        """生成 tkinter Text widget 配置"""
        return {
            'background': self.background,
            'text_color': self.text_color,
            'font_family': self.font_family,
            'font_size': self.font_size,
            'h1_color': self.h1_color,
            'h1_size': self.h1_size,
            'h2_color': self.h2_color,
            'h2_size': self.h2_size,
            'h3_color': self.h3_color,
            'h3_size': self.h3_size,
            'h4_color': self.h4_color,
            'h4_size': self.h4_size,
            'link_color': self.link_color,
            'code_bg': self.code_bg,
            'code_color': self.code_color,
            'code_font': self.code_font,
            'code_block_bg': self.code_block_bg,
            'code_block_color': self.code_block_color,
            'blockquote_bg': self.blockquote_bg,
            'blockquote_color': self.blockquote_color,
            'table_header_bg': self.table_header_bg,
            'table_header_color': self.table_header_color,
        }
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            'name': self.name,
            'display_name': self.display_name,
            'background': self.background,
            'text_color': self.text_color,
            'font_family': self.font_family,
            'font_size': self.font_size,
            'line_height': self.line_height,
            'h1_color': self.h1_color,
            'h1_size': self.h1_size,
            'link_color': self.link_color,
            'code_bg': self.code_bg,
            'code_color': self.code_color,
            'code_block_bg': self.code_block_bg,
            'blockquote_bg': self.blockquote_bg,
            'blockquote_border': self.blockquote_border,
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'PreviewTheme':
        """从字典创建"""
        return cls(**data)


# 预定义主题
GITHUB_THEME = PreviewTheme(
    name='github',
    display_name='GitHub',
    background='#ffffff',
    text_color='#24292f',
    font_family='Segoe UI',
    h1_color='#1f2328',
    h1_border_bottom='1px solid #d0d7de',
    h2_color='#1f2328',
    h2_border_bottom='1px solid #d0d7de',
    link_color='#0969da',
    link_hover_color='#0550ae',
    code_bg='#e8eaed',  # 浅灰色背景（Tkinter不支持透明度）
    code_color='#1f2328',
    code_block_bg='#f6f8fa',
    code_block_color='#24292f',
    blockquote_border='4px solid #d0d7de',
    blockquote_color='#656d76',
)

TYPORA_THEME = PreviewTheme(
    name='typora',
    display_name='Typora',
    background='#ffffff',
    text_color='#333333',
    font_family='Open Sans',
    line_height=1.8,
    h1_color='#333333',
    h1_border_bottom='none',
    h2_color='#333333',
    h2_border_bottom='none',
    link_color='#4183c4',
    code_bg='#f3f4f4',
    code_color='#c7254e',
    code_block_bg='#f8f8f8',
    code_block_color='#333333',
    blockquote_border='4px solid #dfe2e5',
    blockquote_bg='#f9fafb',  # 替换 transparent
)

DARK_THEME = PreviewTheme(
    name='dark',
    display_name='暗色',
    background='#1e1e1e',
    text_color='#d4d4d4',
    font_family='Microsoft YaHei',
    h1_color='#ffffff',
    h1_border_bottom='1px solid #404040',
    h2_color='#e0e0e0',
    h2_border_bottom='1px solid #404040',
    h3_color='#c0c0c0',
    h4_color='#a0a0a0',
    link_color='#569cd6',
    link_hover_color='#9cdcfe',
    code_bg='#2d2d2d',
    code_color='#ce9178',
    code_block_bg='#1e1e1e',
    code_block_color='#d4d4d4',
    blockquote_bg='#2d2d2d',
    blockquote_border='4px solid #569cd6',
    blockquote_color='#9cdcfe',
    table_border='1px solid #404040',
    table_header_bg='#2d2d2d',
    table_header_color='#ffffff',
    table_stripe_bg='#252525',
    hr_color='#404040',
)

NOTION_THEME = PreviewTheme(
    name='notion',
    display_name='Notion',
    background='#ffffff',
    text_color='#37352f',
    font_family='ui-sans-serif',
    font_size=16,
    line_height=1.5,
    h1_color='#37352f',
    h1_size=30,
    h1_border_bottom='none',
    h2_color='#37352f',
    h2_size=24,
    h2_border_bottom='none',
    link_color='#37352f',
    link_decoration='underline',
    code_bg='#e8e7e4',  # 替换 rgba(135,131,120,0.15)
    code_color='#eb5757',
    code_block_bg='#f7f6f3',
    code_block_color='#37352f',
    blockquote_bg='#f9fafb',  # 替换 transparent
    blockquote_border='3px solid #000000',
    blockquote_color='#37352f',
)

ACADEMIC_THEME = PreviewTheme(
    name='academic',
    display_name='学术',
    background='#fffff8',
    text_color='#111111',
    font_family='Georgia',
    font_size=16,
    line_height=1.8,
    h1_color='#111111',
    h1_size=24,
    h1_weight='normal',
    h1_border_bottom='none',
    h2_color='#111111',
    h2_size=20,
    h2_weight='normal',
    h2_border_bottom='none',
    link_color='#0645ad',
    code_bg='#f5f5f5',
    code_color='#333333',
    code_font='Courier New',
    code_block_bg='#f5f5f5',
    code_block_color='#333333',
    blockquote_bg='#f9fafb',  # 替换 transparent
    blockquote_border='2px solid #cccccc',
    blockquote_color='#666666',
)

# 所有预定义主题
BUILTIN_THEMES = {
    'github': GITHUB_THEME,
    'typora': TYPORA_THEME,
    'dark': DARK_THEME,
    'notion': NOTION_THEME,
    'academic': ACADEMIC_THEME,
}


class PreviewThemeManager:
    """预览主题管理器"""
    
    def __init__(self):
        self._themes: Dict[str, PreviewTheme] = dict(BUILTIN_THEMES)
        self._current_theme = 'github'
        self._custom_themes_file = os.path.join(
            os.path.dirname(__file__), 'custom_preview_themes.json'
        )
        self._load_custom_themes()
    
    def _load_custom_themes(self):
        """加载自定义主题"""
        try:
            if os.path.exists(self._custom_themes_file):
                with open(self._custom_themes_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for theme_data in data:
                        theme = PreviewTheme.from_dict(theme_data)
                        self._themes[theme.name] = theme
        except Exception:
            pass
    
    def _save_custom_themes(self):
        """保存自定义主题"""
        try:
            custom_themes = [
                theme.to_dict() for name, theme in self._themes.items()
                if name not in BUILTIN_THEMES
            ]
            with open(self._custom_themes_file, 'w', encoding='utf-8') as f:
                json.dump(custom_themes, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    
    def get_theme(self, name: str) -> Optional[PreviewTheme]:
        """获取主题"""
        return self._themes.get(name)
    
    def get_current_theme(self) -> PreviewTheme:
        """获取当前主题"""
        return self._themes.get(self._current_theme, GITHUB_THEME)
    
    def set_current_theme(self, name: str) -> bool:
        """设置当前主题"""
        if name in self._themes:
            self._current_theme = name
            return True
        return False
    
    def get_all_themes(self) -> Dict[str, PreviewTheme]:
        """获取所有主题"""
        return dict(self._themes)
    
    def get_theme_names(self) -> list:
        """获取所有主题名称"""
        return [(name, theme.display_name) for name, theme in self._themes.items()]
    
    def add_theme(self, theme: PreviewTheme):
        """添加自定义主题"""
        self._themes[theme.name] = theme
        self._save_custom_themes()
    
    def remove_theme(self, name: str) -> bool:
        """删除自定义主题"""
        if name in BUILTIN_THEMES:
            return False  # 不能删除内置主题
        if name in self._themes:
            del self._themes[name]
            self._save_custom_themes()
            return True
        return False
    
    def get_css(self, name: str = None) -> str:
        """获取主题 CSS"""
        theme = self._themes.get(name) if name else self.get_current_theme()
        return theme.to_css() if theme else ""


# 全局主题管理器
preview_theme_manager = PreviewThemeManager()
