# -*- coding: utf-8 -*-

import customtkinter as ctk

from ui.theme import COLORS, COLORS_LIGHT, save_config


class ThemeFeature:
    def __init__(self, app):
        self.app = app

    def apply_mode(self, mode: str):
        ctk.set_appearance_mode("light")
        COLORS.clear()
        COLORS.update(COLORS_LIGHT)

    def toggle_theme(self):
        """单主题模式下不再切换主题。"""
        self.apply_mode("light")
        self.app.config["theme"] = "light"
        save_config(self.app.config)

        try:
            self.app.configure(fg_color=COLORS["bg_light"])
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
            if hasattr(self.app, "input_text"):
                self.app.input_text.container.configure(bg=COLORS["bg_light"])
                self.app.input_text.text_frame.configure(bg=COLORS["bg_light"])
                self.app.input_text.line_numbers.configure(
                    background=COLORS["line_number_bg"],
                    foreground=COLORS["line_number"],
                )
                self.app.input_text.text.configure(
                    bg=COLORS["editor_bg"],
                    fg=COLORS["text_primary"],
                    insertbackground=COLORS["text_primary"],
                )
        except Exception:
            pass

    def update_preview(self):
        try:
            if hasattr(self.app, "preview") and self.app.preview:
                self.app.preview.text.configure(
                    bg=COLORS["preview_bg"],
                    fg=COLORS["text_primary"],
                )
                self.app.preview.text.tag_configure("code", background="#F5F5F5")
                self.app.preview.text.tag_configure("code_block", background="#FAFAFA", foreground="#1F2937")
                self.app.preview.text.tag_configure("link", foreground="#0000FF")
                self.app.preview.text.tag_configure("quote", foreground="#6B7280")
                self.app.preview.text.tag_configure("math", foreground=COLORS["text_primary"])
                self.app.preview.text.tag_configure("math_block", foreground=COLORS["text_primary"])
        except Exception:
            pass

    def update_sidebar(self):
        try:
            if hasattr(self.app, "sidebar") and self.app.sidebar:
                self.app.sidebar.configure(fg_color=COLORS["bg_sidebar"])
        except Exception:
            pass

    def update_cards(self):
        try:
            if hasattr(self.app, "input_card"):
                self.app.input_card.configure(fg_color=COLORS["bg_card"], border_color=COLORS["border"])
            if hasattr(self.app, "preview_card"):
                self.app.preview_card.configure(fg_color=COLORS["bg_card"], border_color=COLORS["border"])
        except Exception:
            pass
