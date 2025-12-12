# -*- coding: utf-8 -*-

import os
import json
import copy
import customtkinter as ctk

# 主题初始化
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.md2word_config.json')

# 亮色主题颜色
COLORS_LIGHT = {
    'primary': '#4F46E5',
    'primary_hover': '#4338CA',
    'secondary': '#EC4899',
    'success': '#10B981',
    'warning': '#F59E0B',
    'danger': '#EF4444',
    'bg_light': '#F9FAFB',
    'bg_card': '#FFFFFF',
    'bg_sidebar': '#F3F4F6',
    'text_primary': '#1E293B',
    'text_secondary': '#6B7280',
    'text_muted': '#94A3B8',
    'border': '#E5E7EB',
    'border_focus': '#4F46E5',
    'line_number': '#94A3B8',
    'line_number_bg': '#F9FAFB',
    'highlight': '#FEF3C7',
    'shadow': '#E5E7EB',
    'editor_bg': '#FFFFFF',
    'preview_bg': '#FFFFFF',
}

# 暗色主题颜色
COLORS_DARK = {
    'primary': '#818CF8',
    'primary_hover': '#6366F1',
    'secondary': '#F472B6',
    'success': '#34D399',
    'warning': '#FBBF24',
    'danger': '#F87171',
    'bg_light': '#1E293B',
    'bg_card': '#334155',
    'bg_sidebar': '#1E293B',
    'text_primary': '#F1F5F9',
    'text_secondary': '#CBD5E1',
    'text_muted': '#94A3B8',
    'border': '#475569',
    'border_focus': '#818CF8',
    'line_number': '#94A3B8',
    'line_number_bg': '#1E293B',
    'highlight': '#854D0E',
    'shadow': '#0F172A',
    'editor_bg': '#1E293B',
    'preview_bg': '#334155',
}

# 当前主题颜色（通过清空+更新的方式在原地修改，供其他模块共享）
COLORS = COLORS_LIGHT.copy()


DEFAULT_EXPORT_STYLE = {
    'body_cn': '宋体',
    'body_en': 'Times New Roman',
    'heading_cn': '黑体',
    'mono': 'Consolas',
    'math': 'Cambria Math',

    'body_size_pt': 12,
    'heading1_size_pt': 22,
    'heading2_size_pt': 16,
    'heading3_size_pt': 15,
    'heading4_size_pt': 14,
    'code_size_pt': 10,
    'caption_size_pt': 10.5,

    'hyperlink_color': '0000FF',
    'hyperlink_underline': True,
    'hyperlink_size_pt': 12,

    'body_alignment': 'left',
    'body_line_spacing': 1.5,
    'body_space_after_pt': 6,
    'body_space_before_pt': 0,
    'body_first_line_indent_pt': 24,

    'heading1_alignment': 'center',
    'heading1_space_before_pt': 24,
    'heading1_space_after_pt': 18,
    'heading1_bold': True,
    'heading2_alignment': 'center',
    'heading2_space_before_pt': 18,
    'heading2_space_after_pt': 12,
    'heading2_bold': True,
    'heading3_alignment': 'left',
    'heading3_space_before_pt': 13,
    'heading3_space_after_pt': 10,
    'heading3_bold': True,
    'heading4_alignment': 'left',
    'heading4_space_before_pt': 10,
    'heading4_space_after_pt': 6,
    'heading4_bold': True,

    'margin_top_cm': 2.54,
    'margin_bottom_cm': 2.54,
    'margin_left_cm': 3.18,
    'margin_right_cm': 3.18,

    'quote_font': 'Times New Roman',
    'quote_size_pt': 12,
    'quote_italic': True,
    'quote_left_indent_cm': 1.0,
    'quote_right_indent_cm': 1.0,
    'quote_space_before_pt': 6,
    'quote_space_after_pt': 6,

    'code_space_before_pt': 6,
    'code_space_after_pt': 6,
    'code_left_indent_cm': 0.5,
    'code_line_spacing': 1.0,

    'image_max_width_in': 6.0,
    'image_caption_position': 'after',
    'image_caption_align': 'center',

    'table_three_line': True,
    'table_alignment': 'center',
    'table_header_bold': True,
    'table_caption_position': 'after',
    'table_caption_align': 'center',

    'image_caption_template': '图 {num}: {text}',
    'table_caption_template': '表 {num}: {text}',
    'caption_font': 'Times New Roman',
}


DEFAULT_CONFIG = {
    'recent_files': [],
    'font_size': 14,
    'theme': 'light',
    'sidebar_visible': True,
    'sidebar_width': 250,
    'export_style': DEFAULT_EXPORT_STYLE,
}


def get_default_export_style() -> dict:
    return copy.deepcopy(DEFAULT_EXPORT_STYLE)


def load_config() -> dict:
    default_config = copy.deepcopy(DEFAULT_CONFIG)
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                merged = {**default_config, **config}
                try:
                    if isinstance(config.get('export_style'), dict):
                        merged['export_style'] = {**default_config.get('export_style', {}), **config.get('export_style', {})}
                except Exception:
                    pass
                return merged
    except (json.JSONDecodeError, IOError, OSError):
        pass
    return default_config


def save_config(config: dict):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except (IOError, OSError):
        pass
