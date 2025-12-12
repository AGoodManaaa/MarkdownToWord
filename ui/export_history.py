# -*- coding: utf-8 -*-

import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import customtkinter as ctk

from ui.theme import COLORS, save_config


def _now_str() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def record_export_event(
    app,
    status: str,
    output_path: Optional[str] = None,
    style: Optional[str] = None,
    page_size: Optional[str] = None,
    error: Optional[str] = None,
) -> None:
    try:
        if not hasattr(app, 'config') or not isinstance(app.config, dict):
            return

        history: List[Dict[str, Any]] = app.config.get('export_history', [])
        if not isinstance(history, list):
            history = []

        item: Dict[str, Any] = {
            'time': _now_str(),
            'status': status,
            'current_file': getattr(app, 'current_file', None),
            'output_path': output_path,
            'style': style,
            'page_size': page_size,
        }
        if error:
            item['error'] = error

        history.insert(0, item)
        app.config['export_history'] = history[:50]
        save_config(app.config)
    except Exception:
        pass


def show_export_history_dialog(app) -> None:
    if not hasattr(app, 'config') or not isinstance(app.config, dict):
        return

    dialog = ctk.CTkToplevel(app)
    dialog.title('导出历史')
    dialog.geometry('720x520')
    dialog.transient(app)
    dialog.grab_set()

    try:
        dialog.update_idletasks()
        x = app.winfo_x() + (app.winfo_width() - 720) // 2
        y = app.winfo_y() + (app.winfo_height() - 520) // 2
        dialog.geometry(f'+{x}+{y}')
    except Exception:
        pass

    container = ctk.CTkFrame(dialog, fg_color=COLORS['bg_card'])
    container.pack(fill='both', expand=True, padx=12, pady=12)

    header = ctk.CTkFrame(container, fg_color='transparent')
    header.pack(fill='x', pady=(4, 8))

    ctk.CTkLabel(
        header,
        text='🕘 导出历史',
        font=ctk.CTkFont(size=16, weight='bold'),
        text_color=COLORS['text_primary'],
    ).pack(side='left')

    def clear_all() -> None:
        try:
            app.config['export_history'] = []
            save_config(app.config)
        except Exception:
            pass
        try:
            dialog.destroy()
        except Exception:
            pass

    ctk.CTkButton(
        header,
        text='清空',
        fg_color='transparent',
        border_width=1,
        border_color=COLORS['border'],
        text_color=COLORS['text_primary'],
        width=80,
        command=clear_all,
    ).pack(side='right')

    scroll = ctk.CTkScrollableFrame(container, fg_color='transparent')
    scroll.pack(fill='both', expand=True)

    history = app.config.get('export_history', [])
    if not isinstance(history, list):
        history = []

    if not history:
        ctk.CTkLabel(
            scroll,
            text='暂无导出记录',
            text_color=COLORS['text_secondary'],
        ).pack(pady=20)
        return

    def open_item(path: Optional[str]) -> None:
        if not path:
            return
        try:
            if os.path.exists(path) and hasattr(app, '_open_file_cross_platform'):
                app._open_file_cross_platform(path)
        except Exception:
            pass

    for item in history:
        row = ctk.CTkFrame(scroll, fg_color=COLORS['bg_light'], corner_radius=10)
        row.pack(fill='x', padx=6, pady=6)

        status = str(item.get('status') or '')
        time_str = str(item.get('time') or '')
        out_path = item.get('output_path')
        cur_file = item.get('current_file')
        style = item.get('style')
        page = item.get('page_size')

        title = f"[{status}] {time_str}"
        subtitle = ''
        if out_path:
            subtitle = os.path.basename(out_path)
        elif cur_file:
            subtitle = os.path.basename(cur_file)

        left = ctk.CTkFrame(row, fg_color='transparent')
        left.pack(side='left', fill='x', expand=True, padx=10, pady=8)

        ctk.CTkLabel(
            left,
            text=title,
            anchor='w',
            text_color=COLORS['text_primary'],
            font=ctk.CTkFont(size=12, weight='bold'),
        ).pack(fill='x')

        meta = f"{subtitle}"
        if style or page:
            meta += f"    ({style or ''} {page or ''})".rstrip()

        ctk.CTkLabel(
            left,
            text=meta,
            anchor='w',
            text_color=COLORS['text_secondary'],
            font=ctk.CTkFont(size=11),
        ).pack(fill='x')

        err = item.get('error')
        if err:
            ctk.CTkLabel(
                left,
                text=str(err)[:200],
                anchor='w',
                text_color=COLORS['danger'],
                font=ctk.CTkFont(size=11),
            ).pack(fill='x', pady=(4, 0))

        btns = ctk.CTkFrame(row, fg_color='transparent')
        btns.pack(side='right', padx=10, pady=10)

        ctk.CTkButton(
            btns,
            text='打开',
            width=70,
            fg_color=COLORS['primary'],
            command=lambda p=out_path: open_item(p),
        ).pack()
