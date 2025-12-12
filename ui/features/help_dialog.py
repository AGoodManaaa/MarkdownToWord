# -*- coding: utf-8 -*-

import customtkinter as ctk

from ui.theme import COLORS


class HelpDialogFeature:
    def __init__(self, app):
        self.app = app

    def show(self):
        """显示快捷键帮助"""
        help_dialog = ctk.CTkToplevel(self.app)
        help_dialog.title("⌨️ 快捷键说明")
        help_dialog.geometry("400x450")
        help_dialog.transient(self.app)
        help_dialog.resizable(False, False)

        # 居中显示
        help_dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 400) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 450) // 2
        help_dialog.geometry(f"+{x}+{y}")

        frame = ctk.CTkFrame(help_dialog, fg_color=COLORS['bg_card'])
        frame.pack(fill='both', expand=True, padx=15, pady=15)

        ctk.CTkLabel(
            frame,
            text="快捷键说明",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COLORS['text_primary'],
        ).pack(pady=(10, 20))

        shortcuts = [
            ("Ctrl+O", "打开文件"),
            ("Ctrl+S", "保存Markdown文件"),
            ("Ctrl+Shift+S", "导出为Word文档"),
            ("Ctrl+Shift+C", "复制到剪贴板"),
            ("Ctrl+K", "命令面板"),
            ("Ctrl+Z", "撤销"),
            ("Ctrl+Y", "重做"),
            ("Ctrl+F", "搜索替换"),
            ("Ctrl+P", "切换预览"),
            ("Ctrl+B", "切换侧边栏"),
            ("Ctrl++/-", "调整字体大小"),
            ("F1", "显示此帮助"),
        ]

        for key, desc in shortcuts:
            row = ctk.CTkFrame(frame, fg_color='transparent')
            row.pack(fill='x', pady=3, padx=20)

            ctk.CTkLabel(
                row,
                text=key,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=COLORS['primary'],
                width=120,
                anchor='w',
            ).pack(side='left')

            ctk.CTkLabel(
                row,
                text=desc,
                font=ctk.CTkFont(size=12),
                text_color=COLORS['text_secondary'],
                anchor='w',
            ).pack(side='left', fill='x', expand=True)

        ctk.CTkButton(
            frame,
            text="确定",
            command=help_dialog.destroy,
            fg_color=COLORS['primary'],
            width=100,
        ).pack(pady=20)
