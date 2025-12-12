# -*- coding: utf-8 -*-

import customtkinter as ctk

from ui.theme import COLORS, COLORS_LIGHT, COLORS_DARK, save_config


class ThemeFeature:
    def __init__(self, app):
        self.app = app

    def apply_mode(self, mode: str):
        mode = "dark" if mode == "dark" else "light"
        ctk.set_appearance_mode(mode)
        COLORS.clear()
        COLORS.update(COLORS_DARK if mode == "dark" else COLORS_LIGHT)

    def toggle_theme(self):
        """切换明暗主题"""
        current = ctk.get_appearance_mode()
        new_mode = "dark" if current == "Light" else "light"

        self.apply_mode(new_mode)

        self.app.config['theme'] = new_mode
        save_config(self.app.config)

        try:
            self.app.theme_btn.configure(text="☀️" if new_mode == "dark" else "🌙")
        except Exception:
            pass

        try:
            self.app.configure(fg_color=COLORS['bg_light'])
        except Exception:
            pass

        self.update_all()

        try:
            self.app.header_styler.update_theme()
        except Exception:
            pass

    def update_all(self):
        self.update_editor()
        self.update_preview()
        self.update_sidebar()
        self.update_cards()

    def update_editor(self):
        try:
            if hasattr(self.app, 'input_text'):
                self.app.input_text.container.configure(bg=COLORS['bg_light'])
                self.app.input_text.text_frame.configure(bg=COLORS['bg_light'])
                self.app.input_text.line_numbers.configure(
                    background=COLORS['line_number_bg'],
                    foreground=COLORS['line_number'],
                )
                self.app.input_text.text.configure(
                    bg=COLORS['editor_bg'],
                    fg=COLORS['text_primary'],
                    insertbackground=COLORS['text_primary'],
                )
        except Exception:
            pass

    def update_preview(self):
        try:
            if hasattr(self.app, 'preview') and self.app.preview:
                self.app.preview.text.configure(
                    bg=COLORS['preview_bg'],
                    fg=COLORS['text_primary'],
                )

                is_dark = ctk.get_appearance_mode() == "Dark"
                code_bg = '#F5F5F5'
                code_block_bg = '#FAFAFA'
                code_block_fg = '#1F2937'
                link_color = '#60A5FA' if is_dark else '#0000FF'
                quote_color = '#9CA3AF' if is_dark else '#6B7280'
                math_color = COLORS['text_primary']

                self.app.preview.text.tag_configure('code', background=code_bg)
                self.app.preview.text.tag_configure('code_block', background=code_block_bg, foreground=code_block_fg)
                self.app.preview.text.tag_configure('link', foreground=link_color)
                self.app.preview.text.tag_configure('quote', foreground=quote_color)
                self.app.preview.text.tag_configure('math', foreground=math_color)
                self.app.preview.text.tag_configure('math_block', foreground=math_color)
        except Exception:
            pass

    def update_sidebar(self):
        try:
            if hasattr(self.app, 'sidebar') and self.app.sidebar:
                self.app.sidebar.configure(fg_color=COLORS['bg_sidebar'])
        except Exception:
            pass

    def update_cards(self):
        try:
            if hasattr(self.app, 'input_card'):
                self.app.input_card.configure(fg_color=COLORS['bg_card'], border_color=COLORS['border'])
            if hasattr(self.app, 'preview_card'):
                self.app.preview_card.configure(fg_color=COLORS['bg_card'], border_color=COLORS['border'])
        except Exception:
            pass
