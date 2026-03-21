# -*- coding: utf-8 -*-

from __future__ import annotations


def insert_example_if_empty_for_app(app) -> None:
    """编辑器为空时自动插入示例内容，并同步欢迎面板显隐。"""
    try:
        current = app.input_text.get("1.0", "end-1c")
        if (current or "").strip():
            app._update_welcome_state()
            return
    except Exception:
        pass

    try:
        app.load_welcome_sample()
    except Exception:
        try:
            app._update_welcome_state()
        except Exception:
            pass
