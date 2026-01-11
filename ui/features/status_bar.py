# -*- coding: utf-8 -*-

import tkinter as tk
from typing import Optional
import customtkinter as ctk

from ui.theme import COLORS


class StatusBarFeature:
    def __init__(self, app):
        self.app = app
        self.frame = None
        self.status_label = None
        self.word_count_label = None
        self.cursor_pos_label = None
        self.progress = None
        self._temp_msg_id = None
        self._pulse_timer = None
        self._pulse_count = 0
        self._default_status = "✨ 就绪 - 支持表格、公式、图片等完整Markdown语法"

    def create(self):
        container = ctk.CTkFrame(self.app, fg_color=COLORS['bg_light'], height=36, corner_radius=0)
        container.pack(fill="x", side="bottom")

        # 顶部分隔线
        sep = ctk.CTkFrame(container, height=1, fg_color=COLORS['border'])
        sep.pack(fill='x', side='top')

        self.frame = ctk.CTkFrame(container, fg_color=COLORS['bg_card'], height=35, corner_radius=0)
        self.frame.pack(fill="x", side="bottom")

        self.status_label = ctk.CTkLabel(
            self.frame,
            text="✨ 就绪 - 支持表格、公式、图片等完整Markdown语法",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary'],
        )
        self.status_label.pack(side="left", padx=18, pady=8)

        self.progress = ctk.CTkProgressBar(
            self.frame,
            width=160,
            height=10,
            corner_radius=10,
            fg_color=COLORS['border'],
            progress_color=COLORS['primary'],
        )
        self.progress.set(0)
        self.progress.pack(side="left", padx=(8, 0), pady=0)
        try:
            self.progress.pack_forget()
        except Exception:
            pass

        right_box = ctk.CTkFrame(self.frame, fg_color="transparent")
        right_box.pack(side="right", padx=18, pady=0)

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

    def update_status(self, message: str, is_temp: bool = False, duration_ms: int = 3000, pulse: bool = False):
        try:
            if self.status_label is not None:
                if self._temp_msg_id:
                    self.app.after_cancel(self._temp_msg_id)
                    self._temp_msg_id = None
                
                self.status_label.configure(text=message)
                
                if pulse:
                    self._start_pulse()
                
                if is_temp:
                    self._temp_msg_id = self.app.after(duration_ms, self._restore_default_status)
        except Exception:
            pass

    def _start_pulse(self):
        """开始呼吸灯效果反馈"""
        if self._pulse_timer:
            self.app.after_cancel(self._pulse_timer)
        self._pulse_count = 0
        self._pulse_step()

    def _pulse_step(self):
        """呼吸灯单步动画"""
        if not self.status_label: return
        
        # 定义颜色序列
        colors = ["#3b82f6", "#60a5fa", "#93c5fd", "#60a5fa", "#3b82f6"]
        
        if self._pulse_count < len(colors) * 2: # 循环两次
            color = colors[self._pulse_count % len(colors)]
            self.status_label.configure(text_color=color)
            self._pulse_count += 1
            self._pulse_timer = self.app.after(150, self._pulse_step)
        else:
            self.status_label.configure(text_color=COLORS['text_secondary'])
            self._pulse_timer = None

    def _restore_default_status(self):
        """恢复默认状态栏信息"""
        try:
            self._temp_msg_id = None
            if self.status_label:
                self.status_label.configure(text=self._default_status)
        except Exception:
            pass

    def update_progress(self, value: Optional[float], text: Optional[str] = None):
        try:
            if text is not None:
                self.update_status(text)
        except Exception:
            pass

        try:
            if self.progress is None:
                return

            if value is None:
                try:
                    self.progress.pack_forget()
                except Exception:
                    pass
                return

            v = float(value)
            if v < 0:
                v = 0.0
            if v > 1:
                v = 1.0

            try:
                if not self.progress.winfo_ismapped():
                    self.progress.pack(side="left", padx=(8, 0), pady=0)
            except Exception:
                pass
            self.progress.set(v)
        except Exception:
            pass

    def update_counts(self, content: str, selected_count: int = 0):
        try:
            if self.word_count_label is None:
                return
            word_count = len((content or "").replace('\n', '').replace(' ', '').replace('\t', ''))
            line_count = (content or "").count('\n') + 1 if content else 0
            paragraphs = [p for p in (content or "").split('\n\n') if p.strip()]
            para_count = len(paragraphs)
            
            text = f"字数: {word_count} | 行数: {line_count} | 段落: {para_count}"
            if selected_count > 0:
                text = f"选中: {selected_count} | " + text
                
            self.word_count_label.configure(text=text)
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
