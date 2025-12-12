# -*- coding: utf-8 -*-

import os
import tempfile

from tkinter import messagebox


class AutoSaveFeature:
    def __init__(self, app, interval_ms: int = 30000):
        self.app = app
        self.interval_ms = interval_ms
        self.after_id = None
        self.file_path = os.path.join(tempfile.gettempdir(), 'md2word_autosave.md')

    def start(self):
        """启动自动保存定时器"""
        self._do_auto_save()

    def _do_auto_save(self):
        """执行自动保存"""
        try:
            content = self.app.input_text.get("1.0", "end-1c")
            if content.strip():
                with open(self.file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
        except Exception:
            pass
        finally:
            try:
                self.after_id = self.app.after(self.interval_ms, self._do_auto_save)
            except Exception:
                self.after_id = None

    def check_recovery(self):
        """检查是否有自动保存的文件可恢复"""
        try:
            if os.path.exists(self.file_path):
                mtime = os.path.getmtime(self.file_path)
                import time
                if time.time() - mtime < 600:
                    with open(self.file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    if content.strip():
                        result = messagebox.askyesno(
                            "恢复自动保存",
                            "发现上次未保存的内容，是否恢复？"
                        )
                        if result:
                            self.app.input_text.delete("1.0", "end")
                            self.app.input_text.insert("1.0", content)
                            self.app.on_text_change(None)
                            self.app.update_status("✅ 已恢复自动保存的内容")
                            return True
                os.remove(self.file_path)
        except Exception:
            pass
        return False

    def clear(self):
        """清除自动保存文件"""
        try:
            if os.path.exists(self.file_path):
                os.remove(self.file_path)
        except Exception:
            pass
