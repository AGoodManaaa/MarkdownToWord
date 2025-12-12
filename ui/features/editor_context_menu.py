# -*- coding: utf-8 -*-

import tkinter as tk


class EditorContextMenuFeature:
    def __init__(self, app):
        self.app = app
        self.menu = None

    def attach(self, text_widget):
        if text_widget is None:
            return
        try:
            self.menu = tk.Menu(text_widget, tearoff=0)
            self.menu.add_command(label="撤销", command=self.app._undo)
            self.menu.add_command(label="重做", command=self.app._redo)
            self.menu.add_separator()
            self.menu.add_command(label="剪切", command=lambda: text_widget.event_generate('<<Cut>>'))
            self.menu.add_command(label="复制", command=lambda: text_widget.event_generate('<<Copy>>'))
            self.menu.add_command(label="粘贴", command=lambda: text_widget.event_generate('<<Paste>>'))
            self.menu.add_separator()
            self.menu.add_command(label="全选", command=lambda: text_widget.event_generate('<<SelectAll>>'))
            self.menu.add_separator()
            self.menu.add_command(label="插入...", command=lambda: self.app.show_insert_menu(None))
            self.menu.add_command(label="导出 Word", command=self.app.export_to_word)
            self.menu.add_command(label="复制到剪贴板", command=self.app.copy_to_clipboard)

            text_widget.bind('<Button-3>', lambda e: self._show(e, text_widget), add='+')
        except Exception:
            pass

    def _show(self, event, text_widget):
        try:
            if self.menu is None:
                return
            text_widget.focus_set()
            try:
                text_widget.mark_set('insert', f"@{event.x},{event.y}")
            except Exception:
                pass
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.menu.grab_release()
            except Exception:
                pass
