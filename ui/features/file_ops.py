# -*- coding: utf-8 -*-

import os

from tkinter import filedialog, messagebox


class FileOpsFeature:
    def __init__(self, app):
        self.app = app

    def open_file(self):
        """打开Markdown文件"""
        if not self.check_unsaved_changes():
            return

        file_path = filedialog.askopenfilename(
            title="选择Markdown文件",
            filetypes=[
                ("Markdown文件", "*.md *.markdown"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )

        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path: str):
        """加载文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", content)
            self.app.current_file = file_path

            self.app._last_saved_content = content
            self.app._content_modified = False
            self.app._update_title()

            self.app.on_text_change(None)

            self.add_recent_file(file_path)

            self.app.update_status(f"✅ 已加载: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件:\n{e}")

    def save_file(self):
        """保存Markdown源文件"""
        if self.app.current_file:
            self.save_to_file(self.app.current_file)
        else:
            self.save_file_as()

    def save_file_as(self):
        """另存为Markdown文件"""
        file_path = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            initialfile="untitled.md",
            filetypes=[
                ("Markdown文件", "*.md"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if file_path:
            self.save_to_file(file_path)
            self.app.current_file = file_path
            self.add_recent_file(file_path)

    def save_to_file(self, file_path: str):
        """实际保存文件"""
        try:
            content = self.app.input_text.get("1.0", "end-1c")
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)

            self.app._last_saved_content = content
            self.app._content_modified = False
            self.app._update_title()
            self.app.update_status(f"✅ 已保存: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存文件:\n{e}")

    def check_unsaved_changes(self) -> bool:
        """检查未保存的更改，返回 True 表示可以继续操作"""
        if not self.app._content_modified:
            return True

        current_content = self.app.input_text.get("1.0", "end-1c")
        if current_content == self.app._last_saved_content:
            return True

        result = messagebox.askyesnocancel(
            "未保存的更改",
            "当前文档有未保存的更改。\n\n是否保存？",
        )

        if result is None:
            return False
        if result:
            self.save_file()
            return True
        return True

    def add_recent_file(self, file_path: str):
        """添加到最近文件列表"""
        recent = self.app.config.get('recent_files', [])

        if file_path in recent:
            recent.remove(file_path)

        recent.insert(0, file_path)

        self.app.config['recent_files'] = recent[:10]
        try:
            from ui.theme import save_config
            save_config(self.app.config)
        except Exception:
            pass

        self.update_recent_files_view()

    def update_recent_files_view(self):
        """更新最近文件视图"""
        if hasattr(self.app, 'recent_files_view'):
            recent = self.app.config.get('recent_files', [])
            self.app.recent_files_view.update_files(recent)

    def open_recent_file(self, file_path: str):
        """打开最近文件"""
        if os.path.exists(file_path):
            self.load_file(file_path)
        else:
            messagebox.showwarning("提示", f"文件不存在:\n{file_path}")
            recent = self.app.config.get('recent_files', [])
            if file_path in recent:
                recent.remove(file_path)
                self.app.config['recent_files'] = recent
                try:
                    from ui.theme import save_config
                    save_config(self.app.config)
                except Exception:
                    pass
                self.update_recent_files_view()
