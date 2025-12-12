# -*- coding: utf-8 -*-

import tkinter as tk
import customtkinter as ctk

from ui.theme import COLORS


class CommandPalette:
    def __init__(self, app):
        self.app = app
        self.win = None
        self.search = None
        self.listbox = None
        self._items = []
        self._filtered = []

    def show(self):
        if self.win is not None and self.win.winfo_exists():
            try:
                self.win.focus()
                return
            except Exception:
                pass

        self.win = ctk.CTkToplevel(self.app)
        self.win.title("命令面板")
        self.win.geometry("520x420")
        self.win.transient(self.app)
        self.win.grab_set()
        self.win.resizable(False, False)

        try:
            self.win.update_idletasks()
            x = self.app.winfo_x() + (self.app.winfo_width() - 520) // 2
            y = self.app.winfo_y() + (self.app.winfo_height() - 420) // 2
            self.win.geometry(f"+{x}+{y}")
        except Exception:
            pass

        container = ctk.CTkFrame(self.win, fg_color=COLORS['bg_card'])
        container.pack(fill='both', expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            container,
            text="⌘ 命令面板",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COLORS['text_primary'],
        ).pack(anchor='w', padx=10, pady=(6, 8))

        self.search = ctk.CTkEntry(
            container,
            placeholder_text="输入关键字…（例如：导出 / 主题 / 预览 / 插入）",
            height=34,
        )
        self.search.pack(fill='x', padx=10)

        list_frame = ctk.CTkFrame(container, fg_color='transparent')
        list_frame.pack(fill='both', expand=True, padx=10, pady=(10, 6))

        self.listbox = tk.Listbox(
            list_frame,
            activestyle='none',
            selectmode='browse',
            bg=COLORS['bg_light'],
            fg=COLORS['text_primary'],
            highlightthickness=1,
            highlightbackground=COLORS['border'],
            selectbackground=COLORS['highlight'],
            selectforeground=COLORS['text_primary'],
            relief='flat',
        )
        self.listbox.pack(side='left', fill='both', expand=True)

        scrollbar = tk.Scrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox.configure(yscrollcommand=scrollbar.set)

        ctk.CTkLabel(
            container,
            text="Enter 执行 | ↑↓ 选择 | Esc 关闭",
            font=ctk.CTkFont(size=12),
            text_color=COLORS['text_secondary'],
        ).pack(anchor='w', padx=10, pady=(0, 6))

        self._items = self._build_items()
        self._filter("")

        self.search.bind('<KeyRelease>', lambda e: self._filter(self.search.get()))
        self.search.bind('<Return>', lambda e: self._run_selected())
        self.search.bind('<Escape>', lambda e: self.close())
        self.search.bind('<Down>', lambda e: self._move(1))
        self.search.bind('<Up>', lambda e: self._move(-1))

        self.listbox.bind('<Double-Button-1>', lambda e: self._run_selected())
        self.listbox.bind('<Return>', lambda e: self._run_selected())
        self.listbox.bind('<Escape>', lambda e: self.close())
        self.listbox.bind('<Down>', lambda e: self._move(1))
        self.listbox.bind('<Up>', lambda e: self._move(-1))

        try:
            self.search.focus()
        except Exception:
            pass

    def close(self):
        try:
            if self.win is not None and self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
        self.win = None

    def _build_items(self):
        a = self.app
        it = getattr(a, 'insert_templates', None)
        return [
            ("打开文件    Ctrl+O", a.open_file),
            ("保存文件    Ctrl+S", a.save_file),
            ("导出 Word    Ctrl+Shift+S", a.export_to_word),
            ("导出样式设置", a.open_export_style_settings),
            ("复制到剪贴板    Ctrl+Shift+C", a.copy_to_clipboard),
            ("搜索/替换    Ctrl+F", a.show_search_dialog),
            ("切换预览    Ctrl+P", a.toggle_preview),
            ("切换侧边栏    Ctrl+B", a.toggle_sidebar),
            ("切换主题", a.toggle_theme),
            ("插入：表格", (it.insert_table_template if it else a.show_insert_menu)),
            ("插入：链接", (it.insert_link_template if it else a.show_insert_menu)),
            ("插入：图片", (it.insert_image_template if it else a.show_insert_menu)),
            ("插入：公式", (it.insert_math_template if it else a.show_insert_menu)),
            ("插入：代码块", (it.insert_code_template if it else a.show_insert_menu)),
            ("插入：任务列表", (it.insert_task_template if it else a.show_insert_menu)),
            ("插入：分割线", (it.insert_hr_template if it else a.show_insert_menu)),
            ("帮助    F1", a.show_help),
        ]

    def _filter(self, query: str):
        q = (query or "").strip().lower()
        filtered = []
        for label, cb in self._items:
            if not q or q in label.lower():
                filtered.append((label, cb))
        self._filtered = filtered

        try:
            self.listbox.delete(0, 'end')
            for label, _ in filtered:
                self.listbox.insert('end', label)
            if filtered:
                self.listbox.selection_clear(0, 'end')
                self.listbox.selection_set(0)
                self.listbox.activate(0)
        except Exception:
            pass

    def _move(self, delta: int):
        try:
            size = self.listbox.size()
            if size <= 0:
                return "break"
            cur = self.listbox.curselection()
            idx = int(cur[0]) if cur else 0
            idx = max(0, min(size - 1, idx + delta))
            self.listbox.selection_clear(0, 'end')
            self.listbox.selection_set(idx)
            self.listbox.activate(idx)
            self.listbox.see(idx)
            return "break"
        except Exception:
            return "break"

    def _run_selected(self):
        try:
            cur = self.listbox.curselection()
            if not cur:
                return
            idx = int(cur[0])
            if idx < 0 or idx >= len(self._filtered):
                return

            label, cb = self._filtered[idx]
            self.close()
            try:
                cb()
                try:
                    self.app.update_status(f"✅ 已执行: {label.split('    ')[0]}")
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass
