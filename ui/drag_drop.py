# -*- coding: utf-8 -*-

from __future__ import annotations

from typing import List

from tkinter import messagebox


def _clean_drop_paths(app, data: str) -> List[str]:
    try:
        parts = list(app.tk.splitlist(data or ''))
    except Exception:
        parts = [data or '']

    clean: List[str] = []
    for p in parts:
        fp = str(p or '').strip()
        if fp.startswith('{') and fp.endswith('}'):
            fp = fp[1:-1]
        if fp:
            clean.append(fp)
    return clean


def handle_drop_for_app(app, event) -> None:
    """处理拖拽导入事件：支持多文件，稳健解析路径。"""
    paths = _clean_drop_paths(app, getattr(event, 'data', '') or '')
    if not paths:
        return

    supported = [p for p in paths if str(p).lower().endswith(('.md', '.markdown', '.txt'))]
    if not supported:
        messagebox.showwarning('提示', '请拖拽Markdown文件(.md, .markdown, .txt)')
        return

    # 打开第一个文件，其余仅加入最近文件
    try:
        app.file_ops.load_file(supported[0])
    except Exception:
        pass

    for fp in supported[1:]:
        try:
            app.file_ops.add_recent_file(fp)
        except Exception:
            pass

    if len(supported) > 1:
        try:
            app.update_status(f'✅ 已导入 {len(supported)} 个文件（已打开第 1 个）')
        except Exception:
            pass
