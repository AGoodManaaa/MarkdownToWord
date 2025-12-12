# -*- coding: utf-8 -*-

from ui.theme import COLORS


class HeaderStyler:
    def __init__(self, app):
        self.app = app

    def style_button(self, btn, active: bool = False):
        try:
            if active:
                btn.configure(
                    fg_color="#FFFFFF",
                    text_color=COLORS['primary'],
                    hover_color="#F3F4F6",
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color="white",
                    hover_color=COLORS['primary_hover'],
                )
        except Exception:
            pass

    def update_theme(self):
        try:
            if hasattr(self.app, 'header'):
                self.app.header.configure(fg_color=COLORS['primary'])
        except Exception:
            pass

        try:
            for btn in getattr(self.app, '_header_default_buttons', []) or []:
                self.style_button(btn, active=False)
        except Exception:
            pass

        try:
            if hasattr(self.app, 'insert_btn') and self.app.insert_btn is not None:
                self.app.insert_btn.configure(fg_color=COLORS['success'], hover_color="#059669", text_color="white")
        except Exception:
            pass

        self.update_states()

    def update_states(self):
        try:
            if hasattr(self.app, 'preview_btn') and self.app.preview_btn is not None:
                self.style_button(self.app.preview_btn, active=bool(getattr(self.app, 'preview_visible', True)))
        except Exception:
            pass

        try:
            if hasattr(self.app, 'sidebar_btn') and self.app.sidebar_btn is not None:
                self.style_button(self.app.sidebar_btn, active=bool(getattr(self.app, 'sidebar_visible', True)))
        except Exception:
            pass
