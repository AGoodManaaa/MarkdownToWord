# -*- coding: utf-8 -*-

import tkinter as tk
import customtkinter as ctk

from ui.theme import COLORS


class StatusBarFeature:
    def __init__(self, app):
        self.app = app
        self.frame = None
        self.status_label = None
        self.word_count_label = None
        self.cursor_pos_label = None

    def create(self):
        self.frame = ctk.CTkFrame(self.app, fg_color=COLORS['bg_card'], height=35, corner_radius=0)
        self.frame.pack(fill="x", side="bottom")

        self.status_label = ctk.CTkLabel(
            self.frame,
            text="✨ 就绪 - 支持表格、公式、图片等完整Markdown语法",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary'],
        )
        self.status_label.pack(side="left", padx=20, pady=8)

        right_box = ctk.CTkFrame(self.frame, fg_color="transparent")
        right_box.pack(side="right", padx=20, pady=0)

        self.word_count_label = ctk.CTkLabel(
            right_box,
            text="字数: 0 | 行数: 0 | 段落: 0",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary'],
            anchor='e',
        )
        self.word_count_label.pack(side="left", pady=8)

        self.cursor_pos_label = ctk.CTkLabel(
            right_box,
            text="行: 1 | 列: 1",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary'],
            anchor='e',
        )
        self.cursor_pos_label.pack(side="left", padx=(12, 0), pady=8)

    def update_status(self, message: str):
        try:
            if self.status_label is not None:
                self.status_label.configure(text=message)
        except Exception:
            pass

    def update_counts(self, content: str):
        try:
            if self.word_count_label is None:
                return
            word_count = len((content or "").replace('\n', '').replace(' ', '').replace('\t', ''))
            line_count = (content or "").count('\n') + 1 if content else 0
            paragraphs = [p for p in (content or "").split('\n\n') if p.strip()]
            para_count = len(paragraphs)
            self.word_count_label.configure(text=f"字数: {word_count} | 行数: {line_count} | 段落: {para_count}")
        except Exception:
            pass

    def update_cursor_position(self, text_widget):
        try:
            if self.cursor_pos_label is None or text_widget is None:
                return
            idx = text_widget.index('insert')
            line_str, col_str = idx.split('.')
            line_no = int(line_str)
            col_no = int(col_str) + 1
            self.cursor_pos_label.configure(text=f"行: {line_no} | 列: {col_no}")
        except Exception:
            pass
