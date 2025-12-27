# -*- coding: utf-8 -*-

import os
import json
import copy
import re
import customtkinter as ctk

# 主题初始化
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.md2word_config.json')

# 亮色主题颜色
COLORS_LIGHT = {
    'primary': '#6366F1',
    'primary_hover': '#4F46E5',
    'secondary': '#F43F5E',
    'success': '#22C55E',
    'warning': '#F59E0B',
    'danger': '#EF4444',
    'bg_light': '#F6F7FB',
    'bg_card': '#FFFFFF',
    'bg_sidebar': '#F0F2F8',
    'text_primary': '#0F172A',
    'text_secondary': '#64748B',
    'text_muted': '#94A3B8',
    'border': '#E6E8F0',
    'border_focus': '#6366F1',
    'line_number': '#94A3B8',
    'line_number_bg': '#F6F7FB',
    'highlight': '#E0E7FF',
    'shadow': '#E6E8F0',
    'editor_bg': '#FFFFFF',
    'preview_bg': '#FFFFFF',
}

# 暗色主题颜色
COLORS_DARK = {
    'primary': '#818CF8',
    'primary_hover': '#6366F1',
    'secondary': '#FB7185',
    'success': '#34D399',
    'warning': '#FBBF24',
    'danger': '#F87171',
    'bg_light': '#0B1220',
    'bg_card': '#0F172A',
    'bg_sidebar': '#0B1220',
    'text_primary': '#E2E8F0',
    'text_secondary': '#94A3B8',
    'text_muted': '#64748B',
    'border': '#1F2A44',
    'border_focus': '#818CF8',
    'line_number': '#64748B',
    'line_number_bg': '#0B1220',
    'highlight': '#1E2A5A',
    'shadow': '#020617',
    'editor_bg': '#0B1220',
    'preview_bg': '#0F172A',
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

    # 样式映射（方案B）：优先使用 Word 样式控制排版
    'use_word_styles': True,
    'map_heading_1': 'Heading 1',
    'map_heading_2': 'Heading 2',
    'map_heading_3': 'Heading 3',
    'map_heading_4': 'Heading 4',
    'map_paragraph': 'Normal',
    'map_quote': 'Block Quote',
    'map_image_caption': 'Caption',
    'map_table_caption': 'Caption',
}


DEFAULT_CONFIG = {
    'recent_files': [],
    'font_size': 14,
    'theme': 'light',
    'sidebar_visible': True,
    'sidebar_width': 250,
    'window_geometries': {},
    'last_open_dir': None,
    'last_save_dir': None,
    'last_export_output_path': None,
    'export_auto_format_markdown': False,
    'export_style': DEFAULT_EXPORT_STYLE,
    'export_style_presets': {},
    'export_history': [],
    'last_export_style': 'standard',
    'last_export_page_size': 'a4',
    'preflight_check_remote_images': False,

    # 导出选项
    'export_toc_enabled': False,
    'export_update_fields_on_open': True,
    
    # 预览缩放
    'preview_zoom_scale': 1.0,
    
    # 多标签页
    'open_tabs': [],
    'active_tab_id': None,
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


def apply_window_icon(win) -> None:
    """统一设置弹窗图标为项目根目录下的 app.ico（Windows）。"""
    try:
        root_dir = os.path.dirname(os.path.dirname(__file__))
        icon_path = os.path.join(root_dir, 'app.ico')
        if not os.path.exists(icon_path):
            return

        def _apply() -> None:
            # Windows: wm_iconbitmap 通常对 Toplevel 更稳定
            try:
                win.wm_iconbitmap(icon_path)
                return
            except Exception:
                pass
            try:
                win.iconbitmap(icon_path)
            except Exception:
                pass

        try:
            # 部分 CTkToplevel 需要等窗口创建后再设置，且可能需要多次重试
            win.after(0, _apply)
            win.after(120, _apply)
            win.after(600, _apply)
        except Exception:
            _apply()
    except Exception:
        pass


def attach_window_geometry(app, win, key: str) -> bool:
    """为任意窗口/弹窗保存并恢复上一次位置与大小。

    - key: 配置中的唯一标识（例如 'export_options' / 'format_dialog'）。
    """
    if not key:
        return False

    restored = False

    # 先恢复
    try:
        cfg = getattr(app, 'config', None)
        if isinstance(cfg, dict):
            store = cfg.get('window_geometries')
            if isinstance(store, dict):
                geo = store.get(str(key))
                if isinstance(geo, str) and geo:
                    try:
                        # 校验 geometry 是否在屏幕范围内（避免跑偏到屏幕外/过低）
                        m = re.match(r'^\s*(\d+)x(\d+)\+(-?\d+)\+(-?\d+)\s*$', geo)
                        if m:
                            w = int(m.group(1))
                            h = int(m.group(2))
                            x = int(m.group(3))
                            y = int(m.group(4))
                            try:
                                sw = int(win.winfo_screenwidth())
                                sh = int(win.winfo_screenheight())
                            except Exception:
                                sw = 0
                                sh = 0

                            # 给一个边界容错，窗口只要至少有一部分可见就认为有效
                            if (sw and sh) and (x > sw - 60 or y > sh - 60 or x < -w + 60 or y < -h + 60):
                                raise ValueError('geometry_out_of_screen')

                        win.geometry(geo)
                        restored = True
                    except Exception:
                        pass
    except Exception:
        pass

    # 再绑定保存（节流）
    state = {'after_id': None}

    def _save_now() -> None:
        try:
            cfg = getattr(app, 'config', None)
            if not isinstance(cfg, dict):
                return
            store = cfg.get('window_geometries')
            if not isinstance(store, dict):
                store = {}
                cfg['window_geometries'] = store

            try:
                geo = str(win.geometry())
            except Exception:
                return
            if not geo:
                return

            store[str(key)] = geo
            try:
                save_config(cfg)
            except Exception:
                pass
        except Exception:
            pass

    def _schedule_save(event=None) -> None:  # noqa: ARG001
        try:
            if state['after_id'] is not None:
                try:
                    win.after_cancel(state['after_id'])
                except Exception:
                    pass
            state['after_id'] = win.after(400, _save_now)
        except Exception:
            state['after_id'] = None

    try:
        win.bind('<Configure>', _schedule_save, add='+')
    except Exception:
        pass

    # 关闭前确保保存一次
    try:
        prev = None
        try:
            prev = win.protocol('WM_DELETE_WINDOW')
        except Exception:
            prev = None

        def _on_close() -> None:
            try:
                _save_now()
            except Exception:
                pass
            try:
                if callable(prev):
                    prev()
                else:
                    win.destroy()
            except Exception:
                try:
                    win.destroy()
                except Exception:
                    pass

        win.protocol('WM_DELETE_WINDOW', _on_close)
    except Exception:
        pass

    return restored
