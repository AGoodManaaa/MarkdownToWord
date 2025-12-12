# -*- coding: utf-8 -*-

import customtkinter as ctk

from ui.theme import COLORS


class InsertTemplatesFeature:
    def __init__(self, app):
        self.app = app
        self._snippet_placeholders = None
        self._snippet_index = -1

    def show_menu(self, event=None):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("插入内容")
        dialog.geometry("500x480")
        dialog.transient(self.app)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 500) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 480) // 2
        dialog.geometry(f"+{x}+{y}")

        ctk.CTkLabel(
            dialog,
            text="➕ 插入内容",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(20, 15))

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="both", expand=True, padx=30, pady=10)

        insert_options = [
            ("📊 表格", "插入三线表样式表格", self.insert_table_template),
            ("🔗 链接", "插入超链接", self.insert_link_template),
            ("🖼️ 图片", "插入图片引用", self.insert_image_template),
            ("π 公式", "插入LaTeX数学公式", self.insert_math_template),
            ("📝 代码块", "插入代码块", self.insert_code_template),
            ("☐ 任务列表", "插入任务清单", self.insert_task_template),
            ("─── 分割线", "插入水平分割线", self.insert_hr_template),
        ]

        for icon_text, desc, cmd in insert_options:
            row = ctk.CTkFrame(btn_frame, fg_color="transparent")
            row.pack(fill="x", pady=6)

            def make_callback(command):
                def callback():
                    dialog.destroy()
                    command()
                return callback

            ctk.CTkButton(
                row,
                text=icon_text,
                font=ctk.CTkFont(size=16),
                width=160,
                height=42,
                fg_color=COLORS['primary'],
                hover_color=COLORS['primary_hover'],
                command=make_callback(cmd),
            ).pack(side="left", padx=(0, 15))

            ctk.CTkLabel(
                row,
                text=desc,
                font=ctk.CTkFont(size=14),
                text_color=COLORS['text_secondary'],
            ).pack(side="left", fill="x")

        ctk.CTkButton(
            dialog,
            text="关闭",
            command=dialog.destroy,
            fg_color="transparent",
            border_width=1,
            border_color=COLORS['border'],
            text_color=COLORS['text_primary'],
            width=100,
        ).pack(pady=20)

    def _insert_template_and_select(self, template: str, select_text=None):
        tb = getattr(self.app.input_text, '_textbox', None)
        if tb is None:
            self.app.insert_text(template)
            return

        insert_idx = tb.index('insert')
        self.app.insert_text(template)

        placeholders = []
        if isinstance(select_text, (list, tuple)):
            placeholders = [p for p in select_text if p]
        elif select_text:
            placeholders = [select_text]

        if not placeholders:
            self._snippet_placeholders = None
            self._snippet_index = -1
            try:
                tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
            except Exception:
                pass
            return

        try:
            start = insert_idx
            end = tb.index(f"{insert_idx}+{len(template)}c")

            tb.mark_set('md2word_snippet_start', start)
            tb.mark_set('md2word_snippet_end', end)
            tb.mark_gravity('md2word_snippet_start', 'left')
            tb.mark_gravity('md2word_snippet_end', 'right')

            selected_idx = None
            selected_range = None
            for i, ph in enumerate(placeholders):
                pos = tb.search(ph, 'md2word_snippet_start', 'md2word_snippet_end')
                if pos:
                    pos_end = tb.index(f"{pos}+{len(ph)}c")
                    selected_idx = i
                    selected_range = (pos, pos_end)
                    break

            if not selected_range:
                self._snippet_placeholders = None
                self._snippet_index = -1
                try:
                    tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
                except Exception:
                    pass
                return

            tb.tag_remove('sel', '1.0', 'end')
            tb.tag_add('sel', selected_range[0], selected_range[1])
            tb.mark_set('insert', selected_range[1])
            tb.see(selected_range[0])
            tb.focus()

            if len(placeholders) > 1 and selected_idx is not None:
                self._snippet_placeholders = placeholders
                self._snippet_index = selected_idx
            else:
                self._snippet_placeholders = None
                self._snippet_index = -1
                try:
                    tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
                except Exception:
                    pass
        except Exception:
            self._snippet_placeholders = None
            self._snippet_index = -1
            try:
                tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
            except Exception:
                pass

    def on_tab(self, event=None):
        tb = getattr(self.app.input_text, '_textbox', None)
        placeholders = getattr(self, '_snippet_placeholders', None)
        if tb is None or not placeholders:
            return None

        try:
            start_bound = tb.index('md2word_snippet_start')
            end_bound = tb.index('md2word_snippet_end')
        except Exception:
            self._snippet_placeholders = None
            self._snippet_index = -1
            return None

        next_index = getattr(self, '_snippet_index', -1) + 1
        if next_index >= len(placeholders):
            self._snippet_placeholders = None
            self._snippet_index = -1
            try:
                tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
            except Exception:
                pass
            return "break"

        for i in range(next_index, len(placeholders)):
            ph = placeholders[i]
            pos = tb.search(ph, start_bound, end_bound)
            if pos:
                pos_end = tb.index(f"{pos}+{len(ph)}c")
                tb.tag_remove('sel', '1.0', 'end')
                tb.tag_add('sel', pos, pos_end)
                tb.mark_set('insert', pos_end)
                tb.see(pos)
                tb.focus()
                self._snippet_index = i
                return "break"

        self._snippet_placeholders = None
        self._snippet_index = -1
        try:
            tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
        except Exception:
            pass
        return "break"

    def on_escape(self, event=None):
        self._snippet_placeholders = None
        self._snippet_index = -1
        tb = getattr(self.app.input_text, '_textbox', None)
        if tb is not None:
            try:
                tb.mark_unset('md2word_snippet_start', 'md2word_snippet_end')
            except Exception:
                pass
        return None

    def insert_table_template(self):
        template = """| 列一 | 列二 | 列三 |\n|------|------|------|\n| 内容1 | 内容2 | 内容3 |\n| 内容4 | 内容5 | 内容6 |\n"""
        self._insert_template_and_select(template, "内容1")
        self.app.update_status("✅ 已插入表格模板")

    def insert_link_template(self):
        template = "[链接文本](https://example.com)"
        self._insert_template_and_select(template, ["链接文本", "https://example.com"])
        self.app.update_status("✅ 已插入链接模板")

    def insert_image_template(self):
        template = "![图片描述](图片路径)"
        self._insert_template_and_select(template, ["图片描述", "图片路径"])
        self.app.update_status("✅ 已插入图片模板")

    def insert_math_template(self):
        template = """$$\n\\frac{a}{b} = c\n$$"""
        self._insert_template_and_select(template, "\\frac{a}{b} = c")
        self.app.update_status("✅ 已插入公式模板")

    def insert_code_template(self):
        template = """```python\n# 在此输入代码\nprint(\"Hello, World!\")\n```"""
        self._insert_template_and_select(template, "# 在此输入代码")
        self.app.update_status("✅ 已插入代码块模板")

    def insert_hr_template(self):
        self.app.insert_text("\n---\n")
        self.app.update_status("✅ 已插入分割线")

    def insert_task_template(self):
        template = """- [ ] 待完成任务 1\n- [ ] 待完成任务 2\n- [x] 已完成任务\n"""
        self._insert_template_and_select(template, "待完成任务 1")
        self.app.update_status("✅ 已插入任务列表模板")
