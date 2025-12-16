# -*- coding: utf-8 -*-

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from ui.theme import COLORS, apply_window_icon, attach_window_geometry
from utils import normalize_markdown, convert_latex_delimiters


def format_markdown_text(md: str) -> str:
    """Markdown 一键规范化：空行/块间距 + LaTeX 分隔符统一。"""
    text = md or ''
    try:
        text = convert_latex_delimiters(text)
    except Exception:
        pass
    try:
        text = normalize_markdown(text)
    except Exception:
        pass
    return text


def show_format_dialog_for_app(app) -> None:
    """展示规范化预览对话框，并允许应用到编辑器。"""
    try:
        content = app.input_text.get('1.0', 'end-1c')
    except Exception:
        content = ''

    if not (content or '').strip():
        messagebox.showwarning('提示', '请先输入Markdown内容')
        return

    formatted = format_markdown_text(content)

    # 跟随主编辑器字号，弹窗内容再略放大一档
    base_size = 14
    try:
        base_size = int((getattr(app, 'config', None) or {}).get('font_size', 14))
    except Exception:
        base_size = 14
    view_size = max(14, base_size)

    dialog = ctk.CTkToplevel(app)
    dialog.title('Markdown 规范化')
    try:
        apply_window_icon(dialog)
    except Exception:
        pass
    restored = False
    try:
        restored = bool(attach_window_geometry(app, dialog, 'format_dialog'))
    except Exception:
        pass
    # 字号变大时，窗口高度也要同步增加，避免底部按钮被挤出
    extra_h = int(max(0, view_size - 14) * 22)
    w, h = 980, 720 + extra_h
    dialog.geometry(f'{w}x{h}')
    try:
        dialog.minsize(900, 660 + extra_h)
        dialog.resizable(True, True)
    except Exception:
        pass
    dialog.transient(app)
    dialog.grab_set()

    # 如果没有恢复到上次位置，则按屏幕居中（避免相对主窗口导致偏下/偏移）
    if not restored:
        try:
            dialog.update_idletasks()
            sw = int(dialog.winfo_screenwidth())
            sh = int(dialog.winfo_screenheight())
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            dialog.geometry(f'+{x}+{y}')
        except Exception:
            pass

    container = ctk.CTkFrame(dialog, fg_color=COLORS['bg_card'])
    container.pack(fill='both', expand=True, padx=14, pady=14)

    ctk.CTkLabel(
        container,
        text='🧹 Markdown 一键规范化',
        font=ctk.CTkFont(size=22, weight='bold'),
        text_color=COLORS['text_primary'],
    ).pack(anchor='w', padx=12, pady=(10, 10))

    tips = '将自动统一空行/块间距，并将 \\(...\\)/\\[...\\] 转为 $...$ / $$...$$。'
    ctk.CTkLabel(
        container,
        text=tips,
        font=ctk.CTkFont(size=14),
        text_color=COLORS['text_secondary'],
        justify='left',
        wraplength=820,
    ).pack(anchor='w', padx=12, pady=(0, 10))

    body = ctk.CTkFrame(container, fg_color='transparent')
    body.pack(fill='both', expand=True, padx=12, pady=(0, 10))
    body.grid_columnconfigure(0, weight=1)
    body.grid_columnconfigure(1, weight=1)
    body.grid_rowconfigure(1, weight=1)

    ctk.CTkLabel(
        body,
        text='原文',
        font=ctk.CTkFont(size=14, weight='bold'),
        text_color=COLORS['text_primary'],
    ).grid(row=0, column=0, sticky='w', pady=(0, 6))

    ctk.CTkLabel(
        body,
        text='规范化后',
        font=ctk.CTkFont(size=14, weight='bold'),
        text_color=COLORS['text_primary'],
    ).grid(row=0, column=1, sticky='w', pady=(0, 6), padx=(12, 0))

    left = tk.Text(body, wrap='word', font=('Segoe UI', view_size))
    left.insert('1.0', content)
    left.configure(state='disabled')
    left.grid(row=1, column=0, sticky='nsew')

    right = tk.Text(body, wrap='word', font=('Segoe UI', view_size))
    right.insert('1.0', formatted)
    right.configure(state='disabled')
    right.grid(row=1, column=1, sticky='nsew', padx=(12, 0))

    btns = ctk.CTkFrame(container, fg_color='transparent')
    btns.pack(fill='x', padx=12, pady=(0, 8))

    def copy_result() -> None:
        try:
            app.clipboard_clear()
            app.clipboard_append(formatted)
            try:
                app.update_status('✅ 已复制规范化结果')
            except Exception:
                pass
        except Exception:
            pass

    def apply_to_editor() -> None:
        try:
            app.input_text.delete('1.0', 'end')
            app.input_text.insert('1.0', formatted)
            try:
                app.on_text_change(None)
            except Exception:
                pass
            try:
                app.update_status('✅ 已应用规范化内容')
            except Exception:
                pass
            dialog.destroy()
        except Exception:
            pass

    ctk.CTkButton(
        btns,
        text='应用到编辑器',
        fg_color=COLORS['primary'],
        font=ctk.CTkFont(size=14, weight='bold'),
        command=apply_to_editor,
        width=140,
    ).pack(side='right')

    ctk.CTkButton(
        btns,
        text='复制规范化结果',
        fg_color='transparent',
        border_width=1,
        border_color=COLORS['border'],
        text_color=COLORS['text_primary'],
        font=ctk.CTkFont(size=13),
        command=copy_result,
        width=150,
    ).pack(side='right', padx=10)

    ctk.CTkButton(
        btns,
        text='关闭',
        fg_color='transparent',
        border_width=1,
        border_color=COLORS['border'],
        text_color=COLORS['text_primary'],
        font=ctk.CTkFont(size=13),
        command=dialog.destroy,
        width=90,
    ).pack(side='left')
