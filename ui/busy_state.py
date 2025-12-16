# -*- coding: utf-8 -*-

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class _WidgetState:
    state: Optional[str] = None


class BusyState:
    """统一管理 App 忙碌态（导出/复制等）。

    目标：
    - 忙碌时禁用关键按钮，避免重复触发。
    - 在状态栏显示明确提示。
    - 导出时允许“取消导出”按钮可用。
    """

    def __init__(self, app: Any):
        self.app = app
        self._busy: bool = False
        self._saved_states: Dict[int, _WidgetState] = {}
        self._reason: Optional[str] = None

    @property
    def is_busy(self) -> bool:
        return self._busy

    def enter(self, reason: str, message: Optional[str] = None) -> None:
        if self._busy:
            return
        self._busy = True
        self._reason = reason

        if message:
            try:
                self.app.update_status(message)
            except Exception:
                pass

        # 需要禁用的控件（尽量覆盖常见入口）
        widgets = []
        for name in (
            'export_btn',
            'copy_btn',
            'clear_btn',
            'export_history_btn',
            'export_style_btn',
            'insert_btn',
            'sidebar_btn',
            'font_minus_btn',
            'font_plus_btn',
            'theme_btn',
            'export_style_header_btn',
        ):
            w = getattr(self.app, name, None)
            if w is not None:
                widgets.append(w)

        # header 默认按钮集合
        try:
            for w in getattr(self.app, '_header_default_buttons', []) or []:
                if w is not None:
                    widgets.append(w)
        except Exception:
            pass

        # 去重并保存状态
        uniq = []
        seen = set()
        for w in widgets:
            wid = id(w)
            if wid in seen:
                continue
            seen.add(wid)
            uniq.append(w)

        for w in uniq:
            try:
                prev = None
                try:
                    prev = str(w.cget('state'))
                except Exception:
                    prev = None
                self._saved_states[id(w)] = _WidgetState(state=prev)
                w.configure(state='disabled')
            except Exception:
                pass

        # 导出时允许取消按钮可用
        if reason == 'export':
            try:
                if hasattr(self.app, 'cancel_export_btn') and self.app.cancel_export_btn is not None:
                    self.app.cancel_export_btn.configure(state='normal')
            except Exception:
                pass

    def exit(self, message: Optional[str] = None) -> None:
        if not self._busy:
            return

        if message:
            try:
                self.app.update_status(message)
            except Exception:
                pass

        # 恢复之前保存的 state
        for name in (
            'export_btn',
            'copy_btn',
            'clear_btn',
            'export_history_btn',
            'export_style_btn',
            'insert_btn',
            'sidebar_btn',
            'font_minus_btn',
            'font_plus_btn',
            'theme_btn',
            'export_style_header_btn',
        ):
            w = getattr(self.app, name, None)
            if w is None:
                continue
            saved = self._saved_states.get(id(w))
            try:
                if saved and saved.state:
                    w.configure(state=saved.state)
                else:
                    w.configure(state='normal')
            except Exception:
                pass

        try:
            for w in getattr(self.app, '_header_default_buttons', []) or []:
                saved = self._saved_states.get(id(w))
                try:
                    if saved and saved.state:
                        w.configure(state=saved.state)
                    else:
                        w.configure(state='normal')
                except Exception:
                    pass
        except Exception:
            pass

        try:
            if hasattr(self.app, 'cancel_export_btn') and self.app.cancel_export_btn is not None:
                self.app.cancel_export_btn.configure(state='disabled')
        except Exception:
            pass

        self._saved_states.clear()
        self._busy = False
        self._reason = None
