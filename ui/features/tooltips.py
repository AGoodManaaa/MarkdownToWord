# -*- coding: utf-8 -*-

import tkinter as tk


class TooltipManager:
    def __init__(self, app, delay_ms: int = 350):
        self.app = app
        self.delay_ms = delay_ms
        self._win = None
        self._after_id = None

    def add_tooltip(self, widget, text: str):
        try:
            widget.bind('<Enter>', lambda e, w=widget, t=text: self._schedule(w, t), add='+')
            widget.bind('<Leave>', lambda e: self.hide(), add='+')
            widget.bind('<ButtonPress-1>', lambda e: self.hide(), add='+')
        except Exception:
            pass

    def _schedule(self, widget, text: str):
        try:
            if self._after_id:
                self.app.after_cancel(self._after_id)
            self._after_id = self.app.after(self.delay_ms, lambda: self.show(widget, text))
        except Exception:
            self._after_id = None

    def show(self, widget, text: str):
        self.hide()
        try:
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 10

            win = tk.Toplevel(self.app)
            win.overrideredirect(True)
            win.attributes('-topmost', True)
            win.configure(bg='#111827')

            label = tk.Label(
                win,
                text=text,
                bg='#111827',
                fg='#F9FAFB',
                padx=10,
                pady=6,
                justify='left',
                font=('Segoe UI', 10),
            )
            label.pack()

            win.update_idletasks()
            w = win.winfo_width()
            h = win.winfo_height()
            win.geometry(f"{w}x{h}+{x - w // 2}+{y}")

            self._win = win
        except Exception:
            self._win = None

    def hide(self):
        try:
            if self._after_id:
                self.app.after_cancel(self._after_id)
        except Exception:
            pass
        self._after_id = None

        try:
            if self._win is not None:
                self._win.destroy()
        except Exception:
            pass
        self._win = None
