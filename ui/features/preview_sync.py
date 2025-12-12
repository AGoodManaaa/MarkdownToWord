# -*- coding: utf-8 -*-

import tkinter as tk


class PreviewSyncFeature:
    def __init__(self, app):
        self.app = app

    def on_text_change_debounced(self, event=None):
        """防抖版文本变化处理 - 300ms延迟"""
        try:
            if getattr(self.app, '_debounce_id', None):
                self.app.after_cancel(self.app._debounce_id)
            self.app._debounce_id = self.app.after(300, lambda: self.on_text_change(event))
        except Exception:
            pass

    def on_text_change(self, event=None):
        """文本变化时更新预览和大纲"""
        content = self.app.input_text.get("1.0", "end-1c")

        # 更新光标位置（即使内容未变化也要更新）
        try:
            self.app._update_cursor_position()
        except Exception:
            pass

        # 内容无变化时跳过预览/大纲重渲染，减少无意义开销
        if getattr(self.app, '_last_content_snapshot', None) == content:
            new_modified = content != getattr(self.app, '_last_saved_content', "")
            if new_modified != getattr(self.app, '_content_modified', False):
                self.app._content_modified = new_modified
                try:
                    self.app._update_title()
                except Exception:
                    pass
            return
        self.app._last_content_snapshot = content

        # 设置预览区为更新状态（防止循环触发）
        if hasattr(self.app, 'preview'):
            try:
                self.app.preview.set_updating(True)
                self.app.preview.update_preview(content)
                self.app.preview.set_updating(False)
            except Exception:
                try:
                    self.app.preview.set_updating(False)
                except Exception:
                    pass

        # 更新大纲
        try:
            if hasattr(self.app, 'outline_view'):
                self.app.outline_view.update_outline(content)
        except Exception:
            pass

        # 更新字数统计
        try:
            self.app.status_bar_feature.update_counts(content)
        except Exception:
            pass

        # 标记内容是否修改
        new_modified = content != getattr(self.app, '_last_saved_content', "")
        if new_modified != getattr(self.app, '_content_modified', False):
            self.app._content_modified = new_modified
            try:
                self.app._update_title()
            except Exception:
                pass

    def on_preview_change(self, markdown_text: str):
        """预览区内容变化时同步回Markdown编辑器"""
        if hasattr(self.app, '_preview_updating') and self.app._preview_updating:
            return

        self.app._preview_updating = True
        try:
            # 保存当前光标位置
            try:
                cursor_pos = self.app.input_text.text.index(tk.INSERT)
            except Exception:
                cursor_pos = None

            self.app.input_text.delete("1.0", "end")
            self.app.input_text.insert("1.0", markdown_text)

            if cursor_pos:
                try:
                    self.app.input_text.text.mark_set(tk.INSERT, cursor_pos)
                except Exception:
                    pass

            self.app.status_bar_feature.update_counts(markdown_text)

            try:
                if hasattr(self.app, 'outline_view'):
                    self.app.outline_view.update_outline(markdown_text)
            except Exception:
                pass

            self.app._last_content_snapshot = markdown_text

            new_modified = markdown_text != getattr(self.app, '_last_saved_content', "")
            if new_modified != getattr(self.app, '_content_modified', False):
                self.app._content_modified = new_modified
                try:
                    self.app._update_title()
                except Exception:
                    pass

            try:
                self.app._update_cursor_position()
            except Exception:
                pass

            try:
                self.app.update_status("✏️ 预览区已编辑")
            except Exception:
                pass
        finally:
            self.app._preview_updating = False

    def on_editor_scroll(self, position: float):
        """编辑器滚动时同步预览区"""
        try:
            if hasattr(self.app, 'preview') and getattr(self.app, 'preview_visible', True):
                self.app.preview.text.yview_moveto(position)
        except Exception:
            pass
