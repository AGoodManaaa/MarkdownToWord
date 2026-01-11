# -*- coding: utf-8 -*-

import tkinter as tk
import time

from ui.features.precise_scroll_sync import PreciseScrollSync, IncrementalPreviewUpdater


class PreviewSyncFeature:
    def __init__(self, app):
        self.app = app

        # 预览/大纲/统计的节流：避免大文档输入时频繁重渲染卡顿
        self._last_preview_ts = 0.0
        self._last_outline_ts = 0.0
        self._last_counts_ts = 0.0
        self._throttle_ms_preview = 120
        self._throttle_ms_outline = 180
        self._throttle_ms_counts = 220
        self._pending_preview_id = None
        
        # 精确滚动同步（延迟初始化）
        self._precise_sync = None
        self._incremental_updater = None
        self._last_content_for_sync = ""
        
        # 行映射更新节流
        self._last_line_map_ts = 0.0
        self._throttle_ms_line_map = 500  # 500ms 节流

        self._scroll_sync_enabled = True

        self._last_editor_scroll_ts = 0.0
        self._last_preview_scroll_ts = 0.0
        self._last_editor_scroll_pos = None
        self._last_preview_scroll_pos = None
        self._scroll_throttle_ms = 80
        self._scroll_delta_threshold = 0.004

    def set_scroll_sync_enabled(self, enabled: bool) -> None:
        try:
            self._scroll_sync_enabled = bool(enabled)
        except Exception:
            self._scroll_sync_enabled = True

        try:
            if hasattr(self.app, 'preview') and hasattr(self.app.preview, 'set_sync_scroll_enabled'):
                self.app.preview.set_sync_scroll_enabled(self._scroll_sync_enabled)
        except Exception:
            pass

    def toggle_scroll_sync(self) -> bool:
        self.set_scroll_sync_enabled(not bool(getattr(self, '_scroll_sync_enabled', True)))
        return bool(getattr(self, '_scroll_sync_enabled', True))

    def _init_precise_sync(self):
        """延迟初始化精确滚动同步"""
        if self._precise_sync is not None:
            return
        
        try:
            if hasattr(self.app, 'input_editor') and hasattr(self.app, 'preview'):
                self._precise_sync = PreciseScrollSync(
                    self.app.input_editor,
                    self.app.preview,
                    self.app
                )
                self._incremental_updater = IncrementalPreviewUpdater(self.app.preview)
        except Exception:
            pass

    def _update_line_map(self, content: str):
        """更新行映射表（带节流）"""
        now = time.monotonic()
        min_interval = float(self._throttle_ms_line_map) / 1000.0
        
        # 节流检查
        if (now - self._last_line_map_ts) < min_interval:
            return
        
        # 内容无变化时跳过
        if content == self._last_content_for_sync:
            return
        
        self._last_line_map_ts = now
        self._last_content_for_sync = content
        
        # 初始化精确同步（如果尚未初始化）
        self._init_precise_sync()
        
        # 更新行映射
        if self._precise_sync:
            try:
                self._precise_sync.build_line_map(content)
            except Exception:
                pass

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
        
        # 更新精确滚动同步的行映射
        self._update_line_map(content)

        now = time.monotonic()

        # 设置预览区为更新状态（防止循环触发）
        # 预览隐藏时直接跳过渲染，减少无意义开销
        if hasattr(self.app, 'preview') and getattr(self.app, 'preview_visible', True):
            try:
                min_interval = float(self._throttle_ms_preview) / 1000.0
                if (now - self._last_preview_ts) >= min_interval:
                    self._last_preview_ts = now
                    try:
                        if self._pending_preview_id:
                            self.app.after_cancel(self._pending_preview_id)
                            self._pending_preview_id = None
                    except Exception:
                        self._pending_preview_id = None

                    self.app.preview.set_updating(True)
                    self.app.preview.update_preview(content)
                    self.app.preview.set_updating(False)
                else:
                    # 合并短时间内的多次更新：只保留最后一次
                    delay = int(max(1, (min_interval - (now - self._last_preview_ts)) * 1000))
                    try:
                        if self._pending_preview_id:
                            self.app.after_cancel(self._pending_preview_id)
                        self._pending_preview_id = self.app.after(
                            delay,
                            lambda c=content: self._render_preview_later(c),
                        )
                    except Exception:
                        pass
            except Exception:
                try:
                    self.app.preview.set_updating(False)
                except Exception:
                    pass

        # 更新大纲（节流）
        try:
            if hasattr(self.app, 'outline_view'):
                min_interval = float(self._throttle_ms_outline) / 1000.0
                if (now - self._last_outline_ts) >= min_interval:
                    self._last_outline_ts = now
                    self.app.outline_view.update_outline(content)
        except Exception:
            pass

        # 更新字数统计（节流）
        try:
            min_interval = float(self._throttle_ms_counts) / 1000.0
            if (now - self._last_counts_ts) >= min_interval:
                self._last_counts_ts = now
                # 使用详细统计功能更新状态栏
                if hasattr(self.app, 'statistics_detail'):
                    self.app.statistics_detail.update_status_bar(content)
                else:
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
            # 更新标签页修改状态
            try:
                if hasattr(self.app, 'tab_manager') and new_modified:
                    self.app.tab_manager.mark_current_modified()
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

    def _render_preview_later(self, content: str) -> None:
        try:
            self._pending_preview_id = None
            if not hasattr(self.app, 'preview'):
                return
            if not getattr(self.app, 'preview_visible', True):
                return

            now = time.monotonic()
            self._last_preview_ts = now
            try:
                self.app.preview.set_updating(True)
                self.app.preview.update_preview(content)
            finally:
                try:
                    self.app.preview.set_updating(False)
                except Exception:
                    pass
        except Exception:
            pass

    def on_editor_scroll(self, position: float):
        """编辑器滚动时同步预览区（使用精确同步）"""
        try:
            if not bool(getattr(self, '_scroll_sync_enabled', True)):
                return
            if not hasattr(self.app, 'preview') or not getattr(self.app, 'preview_visible', True):
                return

            now = time.monotonic()
            if (now - float(getattr(self, '_last_editor_scroll_ts', 0.0))) < (float(self._scroll_throttle_ms) / 1000.0):
                return
            last_pos = getattr(self, '_last_editor_scroll_pos', None)
            if isinstance(last_pos, (int, float)) and abs(float(position) - float(last_pos)) < float(self._scroll_delta_threshold):
                return
            self._last_editor_scroll_ts = now
            self._last_editor_scroll_pos = float(position)
            
            # 初始化精确同步
            self._init_precise_sync()
            
            # 优先使用精确同步
            if self._precise_sync:
                self._precise_sync.sync_editor_to_preview()
            elif hasattr(self.app.preview, 'sync_scroll_to'):
                self.app.preview.sync_scroll_to(position)
            else:
                self.app.preview.text.yview_moveto(position)
        except Exception:
            pass
    
    def on_preview_scroll(self, position: float):
        """预览区滚动时同步编辑器（使用精确同步）"""
        try:
            if not bool(getattr(self, '_scroll_sync_enabled', True)):
                return
            if not hasattr(self.app, 'input_editor'):
                return

            now = time.monotonic()
            if (now - float(getattr(self, '_last_preview_scroll_ts', 0.0))) < (float(self._scroll_throttle_ms) / 1000.0):
                return
            last_pos = getattr(self, '_last_preview_scroll_pos', None)
            if isinstance(last_pos, (int, float)) and abs(float(position) - float(last_pos)) < float(self._scroll_delta_threshold):
                return
            self._last_preview_scroll_ts = now
            self._last_preview_scroll_pos = float(position)
            
            # 初始化精确同步
            self._init_precise_sync()
            
            # 优先使用精确同步
            if self._precise_sync:
                self._precise_sync.sync_preview_to_editor(position)
            elif hasattr(self.app.input_editor, '_textbox'):
                self.app.input_editor._textbox.yview_moveto(position)
        except Exception:
            pass
    
    def get_sync_accuracy(self) -> int:
        """获取当前同步精确度（行数误差）"""
        if not self._precise_sync:
            return -1
        
        try:
            # 获取编辑器当前行
            editor_line = self._precise_sync._get_editor_first_visible_line()
            if editor_line is None:
                return -1
            
            # 获取预览区当前位置对应的行
            if hasattr(self.app.preview, 'text'):
                preview_pos = self.app.preview.text.yview()[0]
                preview_line = self._precise_sync._find_source_line_from_preview(preview_pos)
                if preview_line:
                    return self._precise_sync.get_sync_accuracy(editor_line, preview_line)
            
            return -1
        except Exception:
            return -1
    
    def is_sync_accurate(self, tolerance: int = 2) -> bool:
        """检查同步是否精确（误差在允许范围内）"""
        accuracy = self.get_sync_accuracy()
        return accuracy >= 0 and accuracy <= tolerance
